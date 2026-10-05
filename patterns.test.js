/* patterns.test.js — regression tests for the parser.
 *
 * Run it with:   node patterns.test.js
 *
 * No frameworks, no npm install, no setup. It loads terms.js the same way a
 * browser would and checks each supported pattern still works.
 *
 * Why this exists: the patterns are deliberately loose, and loosening one to
 * catch a missed term is an easy way to start matching things that aren't
 * definitions at all. The NEGATIVE cases below matter as much as the positive
 * ones. Add a case here whenever a new pattern is added or a bug is fixed.
 */
'use strict';

const fs = require('fs');
const path = require('path');

const root = {};
new Function('window', fs.readFileSync(path.join(__dirname, 'terms.js'), 'utf8') +
             '\n;return window;')(root);
const T = root.DefinedTerms;

let passed = 0;
const failures = [];

function check(label, ok, detail) {
  if (ok) { passed++; return; }
  failures.push({ label, detail });
}

/** The term must be indexed from this text. */
function defines(label, term, text) {
  const idx = T.buildIndex(text.split('\n'));
  const got = idx.terms.map(t => t.term);
  const hit = got.some(t => T.keyFor(t) === T.keyFor(term));
  check(label, hit, `expected "${term}", got [${got.join(', ') || 'nothing'}]`);
}

/** Nothing in this text may be indexed — it contains no definitions. */
function definesNothing(label, text) {
  const idx = T.buildIndex(text.split('\n'));
  check(label, idx.terms.length === 0,
        `expected nothing, got [${idx.terms.map(t => t.term).join(', ')}]`);
}

/* ------------------------------------------------------------------ *
 * 1. "X" means …                                                      *
 * ------------------------------------------------------------------ */
defines('plain means', 'Business Day',
  '“Business Day” means any day except a Saturday or Sunday.');
defines('straight quotes', 'EBITDA',
  '"EBITDA" means consolidated net income plus interest expense.');
defines('shall mean', 'Applicable Rate',
  '“Applicable Rate” shall mean, for any day, the rate in the Pricing Grid.');
defines('will mean', 'Facility',
  '“Facility” will mean the revolving credit facility described in Article II.');
defines('means and includes', 'Collateral',
  '“Collateral” means and includes all property subject to the Security Agreement.');
defines('the term X means', 'Knowledge',
  'The term “Knowledge” means the actual knowledge of a Responsible Officer.');
defines('multi-sentence definition kept', 'Consolidated EBITDA',
  '“Consolidated EBITDA” means net income plus interest. For the avoidance of doubt, it excludes gains.');

/* ------------------------------------------------------------------ *
 * 2. Qualifier between the term and the verb, with no comma.          *
 *    This is the bug that hid "Knowledge" — see commit history.       *
 * ------------------------------------------------------------------ */
defines('of qualifier', 'Knowledge',
  '“Knowledge” of the Borrower means the actual knowledge of any Responsible Officer.');
defines('with respect to qualifier', 'Knowledge',
  '“Knowledge” with respect to any Person means the actual knowledge of such Person.');
defines('as used qualifier', 'Knowledge',
  '“Knowledge” as used in this Agreement means the actual knowledge of an officer.');
defines('in relation to qualifier', 'Knowledge',
  '“Knowledge” in relation to any Obligor means the actual knowledge of a director.');
defines('comma clause', 'Knowledge',
  '“Knowledge”, when used with respect to any Person, means the actual knowledge of such Person.');
defines('plural subject verb', 'Knowledge',
  'References to the “Knowledge” of the Borrower mean the actual knowledge of an officer.');

/* ------------------------------------------------------------------ *
 * 3. Cross-references                                                 *
 * ------------------------------------------------------------------ */
defines('has the meaning', 'Disposition',
  '“Disposition” has the meaning assigned to such term in Section 6.05.');
defines('shall have the meanings', 'Permitted Encumbrances',
  '“Permitted Encumbrances” shall have the meanings ascribed thereto in Clause 1.1(b).');

(function pointerIsClean() {
  const idx = T.buildIndex(['“Disposition” has the meaning assigned to such term in Section 6.05.']);
  const e = idx.byKey['disposition'];
  check('pointer captured without trailing stop',
        e && e.pointer === 'Section 6.05',
        `got pointer ${e ? JSON.stringify(e.pointer) : '(no entry)'}`);
})();

/* ------------------------------------------------------------------ *
 * 4. Inline definitions                                               *
 * ------------------------------------------------------------------ */
defines('(the "X")', 'Borrower',
  'This Agreement is entered into among ACME INC. (the “Borrower”) and the lenders.');
defines('(in such capacity, the "X")', 'Administrative Agent',
  'GLOBAL BANK, N.A., as administrative agent (in such capacity, the “Administrative Agent”).');
defines('(collectively, the "X")', 'Secured Parties',
  'The Lenders and the Issuing Bank (collectively, the “Secured Parties”) benefit from the Collateral.');
defines('(a "X")', 'Compliance Certificate',
  'The Borrower shall deliver a certificate of a Financial Officer (a “Compliance Certificate”).');
defines('leading digits allowed', '2026 Notes',
  '“2026 Notes” means the senior notes due 2026 issued by the Borrower.');

/* ------------------------------------------------------------------ *
 * 5. NEGATIVE — ordinary quoted prose must not be indexed             *
 * ------------------------------------------------------------------ */
definesNothing('quoted idiom',
  'The parties agree that “time is of the essence” for each payment obligation.');
definesNothing('scare quotes',
  'No Loan Party shall “double count” any amount in any calculation under this Section.');
definesNothing('reference to another agreement',
  'Reference is made to the “Credit Agreement” dated as of March 1, 2024.');
definesNothing('quoted caption',
  'The rate set forth in the Pricing Grid under the caption “ABR Spread” shall apply.');
definesNothing('quoted sentence',
  'The notice stated “the Borrower has failed to pay the amounts due.”');

/* ------------------------------------------------------------------ *
 * 6. Index behaviour                                                  *
 * ------------------------------------------------------------------ */
(function caseInsensitiveLookup() {
  const idx = T.buildIndex(['“Material Adverse Effect” means a material adverse effect on the business.']);
  check('lookup is case-insensitive',
        !!T.lookup(idx, 'MATERIAL ADVERSE EFFECT') && !!T.lookup(idx, 'material adverse effect'));
})();

(function longestTermWins() {
  const idx = T.buildIndex([
    '“Material Adverse Effect” means a material adverse effect on the business.',
    '“Material” means something of consequence.'
  ]);
  const para = 'A Material Adverse Effect has occurred.';
  const hit = T.matchSelectionText('Adverse', idx, para);
  check('word resolves to the longest containing term',
        hit && hit.term === 'Material Adverse Effect',
        `got ${hit ? hit.term : 'nothing'}`);
})();

(function strongestDefinitionWins() {
  const idx = T.buildIndex([
    'The lender (the “Agent”) shall act for the Lenders.',
    '“Agent” means the institution appointed under Article VIII.'
  ]);
  const e = idx.byKey['agent'];
  check('"means" beats an inline definition',
        e && e.kind === 'means' && e.alternates.length === 1,
        `got kind=${e && e.kind}, alternates=${e && e.alternates.length}`);
})();

/* ------------------------------------------------------------------ *
 * 7. Suspects — the missed-term diagnostic                            *
 * ------------------------------------------------------------------ */
(function suspectsFindUnquoted() {
  const paras = [
    '“Business Day” means any day except a Saturday or Sunday.',
    'Consolidated Net Income shall mean the net income of the Borrower.',
    'The Borrower shall calculate Consolidated Net Income quarterly.'
  ];
  const idx = T.buildIndex(paras);
  const A = T.findSuspects(paras, idx, { minCount: 2 }).filter(s => s.defining);
  check('unquoted definition surfaces in group A',
        A.length === 1 && A[0].phrase === 'Consolidated Net Income',
        `got [${A.map(s => s.phrase).join(', ')}]`);
})();

(function suspectsExcludeKnownTerms() {
  const paras = [
    '“Material Adverse Effect” means a material adverse effect on the business.',
    'No Material Adverse Effect has occurred. A Material Adverse Effect would be bad.'
  ];
  const idx = T.buildIndex(paras);
  const s = T.findSuspects(paras, idx, { minCount: 1 });
  check('indexed terms and their fragments are not suspects',
        s.length === 0, `got [${s.map(x => x.phrase).join(', ')}]`);
})();

(function suspectsExcludePlurals() {
  const paras = [
    '“Subsidiary” means any entity controlled by the Borrower.',
    'The Subsidiaries shall guarantee the Obligations. All Subsidiaries are bound.'
  ];
  const idx = T.buildIndex(paras);
  const s = T.findSuspects(paras, idx, { minCount: 1 }).map(x => x.phrase);
  check('plural of a known term is not a suspect',
        s.indexOf('Subsidiaries') === -1, `got [${s.join(', ')}]`);
})();

(function definingCheckIsAnchoredToThePhrase() {
  // «"EBITDA" shall mean … Consolidated Net Income …» must not be read as a
  // definition OF Consolidated Net Income.
  const paras = [
    '“EBITDA” shall mean Consolidated Net Income plus interest expense.',
    'The Borrower reports Consolidated Net Income annually.'
  ];
  const idx = T.buildIndex(paras);
  const A = T.findSuspects(paras, idx, { minCount: 1 }).filter(s => s.defining);
  check('defining verb must follow the phrase itself',
        A.length === 0, `wrongly flagged [${A.map(s => s.phrase).join(', ')}]`);
})();

/* ------------------------------------------------------------------ */

const total = passed + failures.length;
if (failures.length) {
  console.log(`\n${failures.length} of ${total} FAILED:\n`);
  failures.forEach(f => console.log(`  x ${f.label}\n      ${f.detail}`));
  process.exit(1);
}
console.log(`All ${total} pattern tests passed.`);
