/* app.js — everything that talks to Word.
 *
 * The parsing lives in terms.js. This file only:
 *   1. reads the document's paragraphs and hands them to the parser,
 *   2. listens for the cursor moving and works out which term was clicked,
 *   3. draws the task pane.
 *
 * No network calls are made from this file. Document text never leaves the pane.
 */
(function () {
  'use strict';

  var index = null;          // the current term index
  var activeKey = null;      // which term the pane is showing
  var busy = false;          // a Word.run is in flight
  var queued = false;        // the cursor moved while we were busy
  var selectTimer = null;
  var searchTimer = null;
  var lastError = null;

  var MAX_LISTED = 300;      // cap list rendering so typing stays snappy
  var MAX_CANDIDATES = 18;   // per-paragraph Word searches per click
                             // (a term contributes several surface forms)
  var DEBOUNCE_MS = 180;

  var el = {};

  /* ===================================================================
   * Start-up
   * =================================================================== */

  Office.onReady(function (info) {
    el.status = document.getElementById('status');
    el.refresh = document.getElementById('refresh');
    el.expand = document.getElementById('expand');
    el.follow = document.getElementById('follow');
    el.definition = document.getElementById('definition');
    el.search = document.getElementById('search');
    el.count = document.getElementById('count');
    el.list = document.getElementById('list');
    el.diag = document.getElementById('diag-body');

    if (info.host !== Office.HostType.Word) {
      setStatus('This add-in only runs in Word.');
      return;
    }

    el.refresh.addEventListener('click', function () { scan(); });

    el.expand.addEventListener('click', function () {
      var on = document.body.classList.toggle('expanded');
      el.expand.textContent = on ? 'Show list' : 'Expand';
    });

    el.search.addEventListener('input', function () {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(renderList, 120);
    });

    el.list.addEventListener('click', function (ev) {
      var li = ev.target.closest ? ev.target.closest('li[data-key]') : null;
      if (!li) return;
      showTerm(index && index.byKey[li.getAttribute('data-key')]);
    });

    registerSelectionHandler();
    scan();
  });

  function registerSelectionHandler() {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChanged,
      function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          note('Could not listen for cursor moves: ' + result.error.message);
        }
      }
    );
  }

  /* ===================================================================
   * Scanning the document
   * =================================================================== */

  function scan() {
    if (busy) return;
    busy = true;
    el.refresh.disabled = true;
    setStatus('Scanning…');

    Word.run(function (ctx) {
      var paragraphs = ctx.document.body.paragraphs;
      paragraphs.load('items/text');
      return ctx.sync().then(function () {
        var texts = paragraphs.items.map(function (p) { return p.text; });
        var started = Date.now();
        index = DefinedTerms.buildIndex(texts);
        note('Scanned ' + index.paragraphCount + ' paragraphs in ' +
             (Date.now() - started) + 'ms; ' + index.terms.length + ' terms.');
      });
    })
    .catch(function (err) {
      fail('Scan failed', err);
      index = null;
    })
    .then(function () {
      busy = false;
      el.refresh.disabled = false;
      if (index) {
        setStatus(index.terms.length + ' term' +
                  (index.terms.length === 1 ? '' : 's') + ' indexed');
      }
      activeKey = null;
      renderList();
      renderDefinition(null);
    });
  }

  /* ===================================================================
   * Cursor -> term
   * =================================================================== */

  function onSelectionChanged() {
    if (!el.follow.checked) return;
    clearTimeout(selectTimer);
    selectTimer = setTimeout(function () {
      if (busy) { queued = true; return; }
      resolveAtSelection();
    }, DEBOUNCE_MS);
  }

  function resolveAtSelection() {
    if (!index || !index.terms.length) return;
    busy = true;

    Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.load('text');
      var selParas = sel.paragraphs;
      selParas.load('items/text');

      return ctx.sync().then(function () {
        var paraText = selParas.items.length ? selParas.items[0].text : '';
        var selText = (sel.text || '').replace(/^\s+|\s+$/g, '');

        /* Fast path: the user double-clicked or dragged a selection, so we
         * have actual text to match and never need to ask Word anything. */
        if (selText) {
          var direct = DefinedTerms.matchSelectionText(selText, index, paraText);
          if (direct) return direct;
        }
        if (!paraText) return null;

        /* Caret path. Office.js has no "expand selection to the whole phrase",
         * so we invert the problem: we already know every defined term, so we
         * ask Word to find the ones present in this paragraph and report which
         * of their ranges contains the cursor. Multi-word terms like
         * "Material Adverse Effect" fall out of this for free. */
        /* candidatesIn returns { entry, surface }: the surface is the form the
         * text actually uses, which is what Word must search for. A term
         * defined as "Business Employee(s)" is written "Business Employees". */
        var candidates = DefinedTerms.candidatesIn(paraText, index)
          .filter(searchable)
          .slice(0, MAX_CANDIDATES);
        if (!candidates.length) return null;

        /* These two objects are used again after the next sync. Untracked
         * proxies can go stale between batches, so pin them for the rest of
         * this run; the context is discarded when Word.run returns. */
        var para = selParas.items[0];
        ctx.trackedObjects.add(para);
        ctx.trackedObjects.add(sel);

        var searches = candidates.map(function (c) {
          return {
            entry: c.entry,
            surface: c.surface,
            results: para.search(c.surface, { matchCase: true })
          };
        });
        searches.forEach(function (s) { s.results.load('items/text'); });

        return ctx.sync().then(function () {
          var comparisons = [];
          searches.forEach(function (s) {
            s.results.items.forEach(function (hit) {
              comparisons.push({
                entry: s.entry,
                surface: s.surface,
                relation: hit.compareLocationWith(sel)
              });
            });
          });
          if (!comparisons.length) return null;

          return ctx.sync().then(function () {
            return pickBest(comparisons);
          });
        });
      });
    })
    .then(function (entry) {
      if (entry) showTerm(entry);
    })
    .catch(function (err) {
      /* A click in a header, footer or comment throws rather than returning
       * nothing. That is not worth shouting about, so it goes to Diagnostics. */
      note('Lookup skipped: ' + describeError(err));
    })
    .then(function () {
      busy = false;
      if (queued) { queued = false; resolveAtSelection(); }
    });
  }

  /* How a found range sits relative to the cursor. Lower rank = better match.
   * Word reports these from the found range's point of view, so a cursor
   * sitting inside "Material Adverse Effect" gives Contains. */
  var RELATION_RANK = {
    Equal: 0,
    Contains: 0,
    ContainsStart: 1,
    ContainsEnd: 1,
    Inside: 2,
    InsideStart: 2,
    InsideEnd: 2,
    OverlapsBefore: 3,
    OverlapsAfter: 3,
    AdjacentBefore: 4,   // cursor resting just after the term
    AdjacentAfter: 4     // cursor resting just before it
  };

  function pickBest(comparisons) {
    var best = null;
    for (var i = 0; i < comparisons.length; i++) {
      var c = comparisons[i];
      var rank = RELATION_RANK[c.relation.value];
      if (rank === undefined) continue;
      // Tighter relation wins; ties go to the longer matched text, so clicking
      // inside "Material Adverse Effect" never resolves to "Material".
      if (!best || rank < best.rank ||
          (rank === best.rank && c.surface.length > best.surface.length)) {
        best = { rank: rank, entry: c.entry, surface: c.surface };
      }
    }
    return best ? best.entry : null;
  }

  /* Word's search treats ^ as an escape character and caps the string at 255
   * characters, so skip anything it would choke on. */
  function searchable(candidate) {
    return candidate.surface.length <= 200 && candidate.surface.indexOf('^') === -1;
  }

  /* ===================================================================
   * Rendering
   * =================================================================== */

  function showTerm(entry) {
    if (!entry) return;
    activeKey = entry.key;
    renderDefinition(entry);
    markActiveInList();
  }

  var KIND_LABEL = {
    means: 'means',
    crossref: 'cross-reference',
    inline: 'inline',
    'inline?': 'inline (uncertain)'
  };

  function renderDefinition(entry) {
    if (!entry) {
      if (index && !index.terms.length) {
        el.definition.innerHTML =
          '<div class="empty"><p><strong>No defined terms found.</strong></p>' +
          '<p>This add-in looks for quoted terms, e.g. &ldquo;Material Adverse ' +
          'Effect&rdquo; means&hellip;. If the document defines terms another ' +
          'way, nothing will be indexed.</p></div>';
      } else {
        el.definition.innerHTML =
          '<div class="empty"><p><strong>Click a defined term in the document.</strong></p>' +
          '<p>Its definition appears here. If a single click doesn\'t catch it, ' +
          'double-click the word instead.</p></div>';
      }
      return;
    }

    var weak = entry.confidence < 2 ? ' weak' : '';
    var html =
      '<div class="term-head">' +
        '<span class="term-name">' + esc(entry.term) + '</span>' +
        '<span class="badge' + weak + '">' + esc(KIND_LABEL[entry.kind] || entry.kind) + '</span>' +
      '</div>' +
      '<div class="term-body">' + highlight(entry.text, entry.term) + '</div>';

    var meta = [];
    if (entry.pointer) meta.push('Points to ' + esc(entry.pointer));
    meta.push('paragraph ' + (entry.paragraph + 1));
    html += '<div class="term-meta">' + meta.join(' &middot; ') + '</div>';

    if (entry.alternates && entry.alternates.length) {
      html += '<div class="alt"><div class="alt-head">Also defined ' +
              entry.alternates.length + ' other time' +
              (entry.alternates.length === 1 ? '' : 's') + ':</div>';
      entry.alternates.slice(0, 3).forEach(function (a) {
        html += '<div>&bull; para ' + (a.paragraph + 1) + ' (' +
                esc(KIND_LABEL[a.kind] || a.kind) + '): ' +
                esc(clip(a.text, 160)) + '</div>';
      });
      html += '</div>';
    }

    el.definition.innerHTML = html;
    el.definition.scrollTop = 0;
  }

  function renderList() {
    if (!index) { el.list.innerHTML = ''; el.count.textContent = ''; return; }

    var q = el.search.value.replace(/^\s+|\s+$/g, '').toLowerCase();
    var matches = index.terms.filter(function (t) {
      return !q || t.key.indexOf(q) !== -1;
    });

    var shown = matches.slice(0, MAX_LISTED);
    var html = '';
    for (var i = 0; i < shown.length; i++) {
      var t = shown[i];
      html += '<li data-key="' + esc(t.key) + '"' +
              (t.key === activeKey ? ' class="active"' : '') + '>' +
              '<span class="li-term">' + esc(t.term) + '</span>' +
              '<span class="li-hint">' + esc(clip(stripLead(t.text, t.term), 70)) + '</span>' +
              '</li>';
    }
    el.list.innerHTML = html;

    if (!index.terms.length) {
      el.count.textContent = '';
    } else if (matches.length > shown.length) {
      el.count.textContent = 'Showing ' + shown.length + ' of ' + matches.length + ' matches';
    } else {
      el.count.textContent = matches.length + ' of ' + index.terms.length + ' shown';
    }
  }

  function markActiveInList() {
    var items = el.list.querySelectorAll('li[data-key]');
    for (var i = 0; i < items.length; i++) {
      var on = items[i].getAttribute('data-key') === activeKey;
      items[i].className = on ? 'active' : '';
      if (on && items[i].scrollIntoView) {
        items[i].scrollIntoView({ block: 'nearest' });
      }
    }
  }

  /* ===================================================================
   * Small utilities
   * =================================================================== */

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /* Escape first, then wrap occurrences of the term in <mark>. */
  function highlight(text, term) {
    var safe = esc(text);
    if (!term) return safe;
    var re = new RegExp(DefinedTerms.escapeRegExp(esc(term)), 'g');
    return safe.replace(re, '<mark>$&</mark>');
  }

  /* For the list hint: drop the leading «"Term" means» so the preview shows
   * the substance rather than repeating the term. */
  function stripLead(text, term) {
    var re = new RegExp('^["“”]?' + DefinedTerms.escapeRegExp(term) +
                        '["“”]?\\s*', '');
    return text.replace(re, '');
  }

  function clip(s, n) {
    var t = String(s).replace(/\s+/g, ' ');
    return t.length > n ? t.slice(0, n - 1) + '…' : t;
  }

  function setStatus(text) {
    el.status.textContent = text;
  }

  function describeError(err) {
    if (!err) return 'unknown error';
    var parts = [err.message || String(err)];
    if (err.code) parts.push('code=' + err.code);
    if (err.debugInfo && err.debugInfo.errorLocation) {
      parts.push('at ' + err.debugInfo.errorLocation);
    }
    return parts.join(' | ');
  }

  function fail(what, err) {
    lastError = what + ': ' + describeError(err);
    setStatus(what);
    note(lastError);
  }

  function note(line) {
    var stamp = new Date().toLocaleTimeString();
    el.diag.textContent = stamp + '  ' + line + '\n' + el.diag.textContent;
    if (el.diag.textContent.length > 4000) {
      el.diag.textContent = el.diag.textContent.slice(0, 4000);
    }
  }
})();
