# Where we got to

Last worked on: **5 October 2026**

## Next session starts with UI

**Agreed: UI work comes before desktop deployment and before any of the
"later" features.** Nothing else starts until the UI pass is done.

The UI brief is open — it hadn't been specified when we stopped. Things already
known to be rough:

- The definition card and the term list split the pane 3:2. Workable, but the
  ratio is fixed and can't be dragged.
- **Expand** (hides the list, gives the definition the whole pane) is the only
  way to read a long definition comfortably. It's a blunt instrument.
- Long credit-agreement definitions with (a)/(b)/(c) limbs render as one wall
  of text — no structure, no indentation.
- No way to copy a definition out of the pane.
- No history. Following a chain of terms means re-finding your place by hand.
- The "Also here" chips are the only in-pane navigation that exists.

## What works today

Tested in Word on the web against a real 25,000-word agreement: ~113 terms,
scan is fast, single-click and double-click both resolve.

- Scans the document, indexes defined terms, searchable list, Refresh button.
- Click a term in the document → definition in the pane. **Position decides**,
  so clicking "Business" inside "Business Employees" gives Business
  Employee(s), while a standalone "Business" still gives Business.
- Where a click is ambiguous, the other terms at that spot appear as clickable
  chips ("Also here: …").
- Terms match in every written form: plurals, singulars, possessives, and the
  `Business Employee(s)` bracketed notation.
- Patterns: `"X" means` / `shall mean`, `(the "X")` inline, `"X" has the
  meaning given in Clause Y`, and qualifiers before the verb
  (`"Knowledge" of the Borrower means …`).

## Parked, by agreement

Artifacts in the test bench, judged ignorable for now:

1. **Cross-referenced terms appear under "quoted and not indexed".** This one
   is the table *lying* — the term IS indexed. It tests each quoted span in
   isolation instead of asking whether the term reached the index. Two-line
   fix: skip spans whose text resolves to a known term. Worth doing before
   trusting that table again.
2. Section headings showing in Group B. Correct, just unhelpful.
3. `Seller's Knowledge` in Group B — mid-phrase possessive isn't stripped, so
   it doesn't reduce to Seller + Knowledge. Clicking **Knowledge** in that
   phrase already works; this is diagnostic noise only.

## Still unproven

- **Windows desktop has never been tested.** Everything so far is Word on the
  web. README Step 4 covers the shared-folder sideload.
- Unquoted definitions (`Consolidated Net Income shall mean …`) are not
  indexed by design — they surface in Group A instead. No evidence yet from
  real documents that the pattern is needed; decide when some appears.
- Definitions in headers, footers and footnotes are invisible to
  `body.paragraphs`. Affects both the add-in and the test bench.

## Later features, not started

Clickable nested terms inside a definition · go-to-definition · following
`has the meaning given in Clause X` through to the text · flagging duplicate /
unused / undefined terms.

Groundwork exists: every entry records its paragraph number, any clause
pointer, and competing definitions found elsewhere (`alternates`).

## Working on this

```
node patterns.test.js     55 tests, under a second, no setup
test.html                 paste a document, tune patterns, no Word needed
```

Run the tests before and after any change to `terms.js`. Roughly half are
*negative* cases — quoted prose that must NOT be indexed — because the way
loose patterns fail is by matching ordinary quotations.

Two bugs found so far had the same shape, and it's worth watching for a third:
**the parser and the diagnostic sharing a blind spot**, so the tool that is
supposed to find the gap cannot see it either. That is why `definesPhrase()`
and the parser now share one regex fragment.

## Deploy

GitHub Pages serves <https://parkerland.github.io/Defined/> from `main`.

- After changing `app.js` / `terms.js` / `taskpane.css`, bump `?v=N` in
  `index.html` (and `test.html` for `terms.js`), or Word serves the old file.
- Pages builds are usually under a minute but have taken up to nine. If the
  pane looks unchanged, check the build finished before suspecting the code.
- Stale after a successful build? `Ctrl+Shift+R` on the web; on desktop delete
  everything in `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` with Word closed.

**State at stop:** commit `71f9810` pushed; its Pages build was still running.
Verify the live site serves `?v=6` and that `terms.js` contains
`resolveOverlaps` before concluding anything is broken.
