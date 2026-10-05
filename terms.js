/* terms.js — defined-term parsing for the "Defined" Word add-in.
 *
 * This file is deliberately PURE: it knows nothing about Word, Office.js or the
 * DOM. It takes an array of paragraph strings and returns an index. That means
 * you can open test.html in any browser, paste in a real agreement, and tune the
 * patterns without going near Word.
 *
 * Written in conservative ES5-style JavaScript (no ?., no ??, no \p{...}) so it
 * runs in every webview Word might host the task pane in.
 */
(function (root) {
  'use strict';

  /* ---------------------------------------------------------------------
   * Quote handling
   *
   * Real documents from EDGAR use curly quotes (“X”), not straight ones ("X").
   * Mixing the two is the single most common reason a first attempt at this
   * finds nothing, so every quote character is handled from the start.
   * ------------------------------------------------------------------- */
  var QUOTE_CHARS = '"“”„‟«»';
  var Q = '[' + QUOTE_CHARS + ']';
  var NOT_Q = '[^' + QUOTE_CHARS + '\\r\\n]';

  /* Every quoted span in a paragraph. Lazy inner match so "A" and "B" yields
   * two spans rather than one span swallowing the middle. */
  var QUOTED_RE = new RegExp(Q + '(' + NOT_Q + '{1,120}?)' + Q, 'g');

  /* Cheap pre-filter: skip paragraphs with no quote mark at all. */
  var HAS_QUOTE_RE = new RegExp(Q);

  /* What may sit between the closing quote and the defining verb:
   * an aside in brackets, and/or a comma clause ("X", when used in this
   * Agreement, means ...). Kept free of sentence-ending punctuation so it
   * cannot run across into the next sentence. */
  var LEAD = '^(?:\\s*\\([^)\\r\\n]{0,80}\\))?(?:\\s*,[^.;!?\\r\\n]{0,160})?\\s*[,:\\u2013\\u2014-]?\\s*';

  /* "X" has the meaning given in Clause 1.1 — tested BEFORE the means pattern,
   * because "shall have the meaning" would otherwise be caught by it. */
  var CROSSREF_RE = new RegExp(
    LEAD + '(?:shall\\s+have|has|have|shall\\s+bear|bears)\\s+the\\s+(?:respective\\s+)?meanings?\\b', 'i');

  /* "X" means / shall mean / will mean / means and includes */
  var MEANS_RE = new RegExp(
    LEAD + '(?:means\\s+and\\s+includes|means|shall\\s+mean|will\\s+mean|shall\\s+be\\s+construed\\s+as)\\b', 'i');

  /* The clause a cross-reference points at. Captured now, followed later. */
  var POINTER_RE = new RegExp(
    '\\b(Clause|Section|Article|Schedule|Exhibit|Annex|Appendix|Paragraph|Part)\\s+' +
    '([0-9A-Za-z][\\w.()\\u2013-]*)', 'i');

  /* An inline definition: (the "X"), ("X"), (each, a "X"), (collectively, the "X"),
   * (hereinafter referred to as the "X"). Matched against the text immediately
   * BEFORE the quote, anchored at its end.
   *
   * Repetition is bounded ({0,8}) rather than open-ended: nested * quantifiers
   * on overlapping character classes can backtrack catastrophically. */
  var INLINE_STRICT_RE = new RegExp(
    '\\((?:\\s{0,4}(?:' +
      // "in such capacity, the "Administrative Agent"" — ubiquitous in credit
      // agreements, so these phrases are matched explicitly rather than by
      // loosening the word list (which would cost precision).
      'in\\s{1,2}(?:such|its|that|the\\s{1,2}same)\\s{1,2}capacit(?:y|ies)|' +
      'acting\\s{1,2}in\\s{1,2}such\\s{1,2}capacity|' +
      'as\\s{1,2}(?:hereinafter\\s{1,2})?defined(?:\\s{1,2}(?:below|above|herein))?|' +
      'referred\\s{1,2}to\\s{1,2}as|known\\s{1,2}as|called|defined\\s{1,2}as|' +
      'in\\s{1,2}each\\s{1,2}case|as\\s{1,2}applicable|whether\\s{1,2}one\\s{1,2}or\\s{1,2}more|' +
      'the|each|a|an|any|all|collectively|together|individually|respectively|' +
      'and|or|being|such|this|these|herein|hereinafter|hereafter' +
    ')[\\s,]{0,4}){0,8}$', 'i');

  /* A looser fallback: any open bracket close behind, e.g. (as defined below, the "X").
   * Indexed with lower confidence so you can tell the two apart in the UI. */
  var INLINE_LOOSE_RE = new RegExp('\\([^()\\r\\n]{0,60}$');

  /* Abbreviations that end in a full stop but do not end a sentence. */
  var ABBREVIATIONS = ['no', 'inc', 'ltd', 'co', 'corp', 'plc', 'llc', 'lp', 'llp',
    'art', 'sec', 'cl', 'para', 'pp', 'vol', 'eg', 'ie', 'etc', 'viz', 'cf',
    'jan', 'feb', 'mar', 'apr', 'jun', 'jul', 'aug', 'sep', 'sept', 'oct', 'nov', 'dec'];

  var MAX_DEFINITION_CHARS = 2500;

  var CONFIDENCE = { means: 3, crossref: 2, inline: 2, 'inline?': 1 };

  /* --------------------------------------------------------------------- */
  /* Small helpers                                                          */
  /* --------------------------------------------------------------------- */

  function trim(s) {
    return String(s).replace(/^[\s ]+|[\s ]+$/g, '');
  }

  function collapseSpace(s) {
    return trim(String(s).replace(/[\s ]+/g, ' '));
  }

  /* The lookup key for a term: case- and punctuation-insensitive, so
   * “Material Adverse Effect” and “MATERIAL ADVERSE EFFECT” are one entry. */
  function keyFor(term) {
    return collapseSpace(String(term).toLowerCase())
      .replace(new RegExp(Q, 'g'), '')
      .replace(/[‘’']/g, "'")
      .replace(/[–—]/g, '-');
  }

  function escapeRegExp(s) {
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /* Is this plausibly a defined term, rather than an ordinary quotation?
   * Quoted text is common in these documents ("the Borrower shall not...") so
   * this filter is what keeps the index clean. */
  function isPlausibleTerm(term, kind) {
    var t = collapseSpace(term);
    if (!t || t.length > 120) return false;
    if (!/[A-Za-z]/.test(t)) return false;          // must contain a letter
    if (/[.!?]$/.test(t)) return false;             // a sentence, not a term
    if (t.split(' ').length > 10) return false;     // too long to be a term
    if (/[;]/.test(t)) return false;

    // Inline definitions are the noisiest pattern, so require the term to look
    // like a proper noun: first letter capitalised, or an acronym, or "2026 Notes".
    if (kind === 'inline' || kind === 'inline?') {
      if (!/^[^A-Za-z0-9]*[A-Z0-9]/.test(t)) return false;
    }
    return true;
  }

  function endsWithAbbreviation(text, dotIndex) {
    var before = text.slice(Math.max(0, dotIndex - 12), dotIndex);
    var m = before.match(/([A-Za-z]+)$/);
    if (!m) return false;
    var word = m[1].toLowerCase();
    if (word.length === 1) return true;                       // initials: J. Smith, U.S.
    return ABBREVIATIONS.indexOf(word) !== -1;
  }

  /* Index of the first character of the sentence containing position i. */
  function sentenceStart(text, i) {
    for (var p = Math.min(i, text.length - 1); p > 0; p--) {
      var c = text.charAt(p);
      if (c === '.' || c === '!' || c === '?') {
        if (endsWithAbbreviation(text, p)) continue;
        var rest = text.slice(p + 1);
        var m = rest.match(/^\s+/);
        if (m) return p + 1 + m[0].length;
      }
    }
    return 0;
  }

  /* Index just past the end of the sentence containing position i. */
  function sentenceEnd(text, i) {
    for (var p = i; p < text.length; p++) {
      var c = text.charAt(p);
      if (c === '.' || c === '!' || c === '?') {
        if (endsWithAbbreviation(text, p)) continue;
        var after = text.slice(p + 1, p + 3);
        if (after === '' || /^\s/.test(after)) return p + 1;
      }
    }
    return text.length;
  }

  /* Definitions in a definitions schedule are usually a whole paragraph, which
   * is exactly what you want to read. Only trim if it is genuinely enormous. */
  function capAtSentence(text) {
    if (text.length <= MAX_DEFINITION_CHARS) return text;
    var cut = text.slice(0, MAX_DEFINITION_CHARS);
    var lastStop = Math.max(cut.lastIndexOf('. '), cut.lastIndexOf('; '));
    if (lastStop > MAX_DEFINITION_CHARS * 0.5) cut = cut.slice(0, lastStop + 1);
    return trim(cut) + ' …';
  }

  /* --------------------------------------------------------------------- */
  /* Classification                                                         */
  /* --------------------------------------------------------------------- */

  /* Given one quoted span, decide whether it is a definition and of what kind.
   * Returns null for ordinary quoted text. */
  function classify(paraText, span) {
    var after = paraText.slice(span.end);
    var before = paraText.slice(Math.max(0, span.start - 80), span.start);

    var crossref = CROSSREF_RE.exec(after);
    if (crossref) {
      var tail = after.slice(0, 400);
      var ptr = POINTER_RE.exec(tail);
      return {
        kind: 'crossref',
        // Strip the sentence's full stop off "Clause 1.1(b)." etc.
        pointer: ptr ? collapseSpace(ptr[0]).replace(/[.,;:]+$/, '') : null,
        text: capAtSentence(trim(paraText.slice(span.start, sentenceEnd(paraText, span.end))))
      };
    }

    if (MEANS_RE.test(after)) {
      // Run to the end of the paragraph: in a definitions clause the paragraph
      // IS the definition, including any (a)/(b) limbs and provisos.
      return {
        kind: 'means',
        pointer: null,
        text: capAtSentence(trim(paraText.slice(span.start)))
      };
    }

    var inlineKind = null;
    if (INLINE_STRICT_RE.test(before)) inlineKind = 'inline';
    else if (INLINE_LOOSE_RE.test(before)) inlineKind = 'inline?';

    if (inlineKind) {
      // The useful text is the sentence that introduces the term, not the
      // bracket itself.
      var s = sentenceStart(paraText, span.start);
      var e = sentenceEnd(paraText, span.end);
      return {
        kind: inlineKind,
        pointer: null,
        text: capAtSentence(trim(paraText.slice(s, e)))
      };
    }

    return null;
  }

  /* --------------------------------------------------------------------- */
  /* Index building                                                         */
  /* --------------------------------------------------------------------- */

  /**
   * Build the term index.
   * @param {string[]} paragraphs - the document, one string per paragraph.
   * @returns {object} index
   */
  function buildIndex(paragraphs) {
    var list = paragraphs || [];
    var byKey = {};
    var scanned = 0;

    for (var i = 0; i < list.length; i++) {
      var paraText = String(list[i] == null ? '' : list[i]).replace(/\r/g, '');
      if (!paraText || !HAS_QUOTE_RE.test(paraText)) continue;
      scanned++;

      QUOTED_RE.lastIndex = 0;
      var m;
      while ((m = QUOTED_RE.exec(paraText)) !== null) {
        var raw = m[1];
        var span = { start: m.index, end: m.index + m[0].length };
        var hit = classify(paraText, span);
        if (!hit) continue;

        var term = collapseSpace(raw);
        if (!isPlausibleTerm(term, hit.kind)) continue;

        addEntry(byKey, {
          term: term,
          key: keyFor(term),
          kind: hit.kind,
          confidence: CONFIDENCE[hit.kind] || 1,
          pointer: hit.pointer,
          text: hit.text,
          paragraph: i
        });
      }
    }

    var terms = [];
    for (var k in byKey) {
      if (Object.prototype.hasOwnProperty.call(byKey, k)) terms.push(byKey[k]);
    }
    terms.sort(function (a, b) { return a.key < b.key ? -1 : a.key > b.key ? 1 : 0; });

    return {
      terms: terms,
      byKey: byKey,
      paragraphCount: list.length,
      paragraphsWithQuotes: scanned,
      scannedAt: new Date()
    };
  }

  /* Merge a new hit into the index. The strongest definition wins; weaker ones
   * are kept as alternates (groundwork for duplicate-definition flagging). */
  function addEntry(byKey, entry) {
    var existing = byKey[entry.key];
    if (!existing) {
      entry.alternates = [];
      byKey[entry.key] = entry;
      return;
    }
    var better =
      entry.confidence > existing.confidence ||
      (entry.confidence === existing.confidence && entry.text.length > existing.text.length);

    if (better) {
      entry.alternates = existing.alternates.concat([{
        kind: existing.kind, text: existing.text,
        paragraph: existing.paragraph, pointer: existing.pointer
      }]);
      byKey[entry.key] = entry;
    } else {
      existing.alternates.push({
        kind: entry.kind, text: entry.text,
        paragraph: entry.paragraph, pointer: entry.pointer
      });
    }
  }

  /* --------------------------------------------------------------------- */
  /* Lookup helpers — used by the click-to-define path in app.js             */
  /* --------------------------------------------------------------------- */

  function lookup(index, text) {
    if (!index) return null;
    var e = index.byKey[keyFor(text)];
    return e || null;
  }

  /**
   * Which indexed terms literally appear in this paragraph?
   * Longest first, so "Material Adverse Effect" beats "Material" when both
   * ranges contain the caret.
   */
  function candidatesIn(paraText, index) {
    if (!index || !paraText) return [];
    var out = [];
    for (var i = 0; i < index.terms.length; i++) {
      var t = index.terms[i];
      if (paraText.indexOf(t.term) !== -1) out.push(t);
    }
    out.sort(function (a, b) { return b.term.length - a.term.length; });
    return out;
  }

  /**
   * Resolve selected text (a double-click, or a dragged selection) to a term.
   * Tries exact match, then the longest indexed term in this paragraph that
   * contains the selection, then the longest term inside the selection.
   */
  function matchSelectionText(selText, index, paraText) {
    var sel = collapseSpace(selText);
    if (!sel || !index) return null;

    var exact = lookup(index, sel);
    if (exact) return exact;

    var pool = paraText ? candidatesIn(paraText, index) : index.terms;
    var word = new RegExp('(^|[^A-Za-z0-9])' + escapeRegExp(sel) + '([^A-Za-z0-9]|$)', 'i');
    var i;

    // The selection is one word of a longer term: "Adverse" -> "Material Adverse Effect"
    var containing = [];
    for (i = 0; i < pool.length; i++) {
      if (word.test(pool[i].term)) containing.push(pool[i]);
    }
    if (containing.length) {
      containing.sort(function (a, b) { return b.term.length - a.term.length; });
      return containing[0];
    }

    // The selection spans more than the term: pick the longest term inside it.
    var inside = [];
    for (i = 0; i < pool.length; i++) {
      if (sel.indexOf(pool[i].term) !== -1) inside.push(pool[i]);
    }
    if (inside.length) {
      inside.sort(function (a, b) { return b.term.length - a.term.length; });
      return inside[0];
    }
    return null;
  }

  root.DefinedTerms = {
    buildIndex: buildIndex,
    lookup: lookup,
    candidatesIn: candidatesIn,
    matchSelectionText: matchSelectionText,
    keyFor: keyFor,
    escapeRegExp: escapeRegExp,
    _internals: {
      classify: classify,
      isPlausibleTerm: isPlausibleTerm,
      sentenceStart: sentenceStart,
      sentenceEnd: sentenceEnd,
      QUOTED_RE: QUOTED_RE
    }
  };
})(typeof window !== 'undefined' ? window : this);
