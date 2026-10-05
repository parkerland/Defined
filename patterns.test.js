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
 * 6b. Surface forms — a term is defined once but written many ways    *
 * ------------------------------------------------------------------ */
(function bracketedPluralNotation() {
  const paras = [
    '“Business Employee(s)” means each employee of the Business listed on Schedule 2.',
    'The Buyer shall offer employment to the Business Employees on the Closing Date.'
  ];
  const idx = T.buildIndex(paras);
  const e = idx.byKey['business employee(s)'];
  check('"Business Employee(s)" is indexed', !!e);

  check('clicking "Business Employees" resolves to it',
        (T.matchSelectionText('Business Employees', idx, paras[1]) || {}).term === 'Business Employee(s)',
        `got ${JSON.stringify((T.matchSelectionText('Business Employees', idx, paras[1]) || {}).term)}`);
  check('clicking the singular resolves too',
        (T.matchSelectionText('Business Employee', idx, paras[1]) || {}).term === 'Business Employee(s)');

  const surfaces = T.candidatesIn(paras[1], idx).map(c => c.surface);
  check('the plural is offered as a searchable surface form',
        surfaces.indexOf('Business Employees') !== -1, `got [${surfaces.join(', ')}]`);

  const sus = T.findSuspects(paras, idx, { minCount: 1 }).map(s => s.phrase);
  check('the plural is not reported as an unknown phrase',
        sus.indexOf('Business Employees') === -1, `got [${sus.join(', ')}]`);
})();

(function ordinaryPlurals() {
  const idx = T.buildIndex([
    '“Subsidiary” means any entity controlled by the Borrower.',
    '“Party” means a party to this Agreement.'
  ]);
  const para = 'The Subsidiaries and the Parties shall comply.';
  check('Subsidiaries -> Subsidiary',
        (T.matchSelectionText('Subsidiaries', idx, para) || {}).term === 'Subsidiary');
  check('Parties -> Party',
        (T.matchSelectionText('Parties', idx, para) || {}).term === 'Party');
})();

(function singularOfAPluralTerm() {
  const idx = T.buildIndex(['“Loan Documents” means this Agreement and each Note.']);
  check('Loan Document -> Loan Documents',
        (T.matchSelectionText('Loan Document', idx, 'Each Loan Document is binding.') || {}).term
          === 'Loan Documents');
})();

(function possessiveIsStripped() {
  const idx = T.buildIndex(['“Borrower” means ACME INC., a Delaware corporation.']);
  check("Borrower's -> Borrower",
        (T.matchSelectionText('Borrower’s', idx, 'The Borrower’s obligations.') || {}).term
          === 'Borrower');
})();

(function distinctTermsKeepTheirOwnEntries() {
  // "Lender" and "Lenders" defined separately must not collapse into one.
  const idx = T.buildIndex([
    '“Lender” means each financial institution party hereto.',
    '“Lenders” means all of the Lender parties collectively.'
  ]);
  check('a variant never shadows another term\'s own definition',
        (T.lookup(idx, 'Lender') || {}).term === 'Lender' &&
        (T.lookup(idx, 'Lenders') || {}).term === 'Lenders');
})();

/* ------------------------------------------------------------------ *
 * 6c. Clicking where terms overlap                                     *
 *                                                                      *
 * "Business", "Business Day" and "Business Employee(s)" can all be     *
 * defined, and the word "Business" belongs to all three. Which one you *
 * meant depends on WHERE you clicked. Word supplies the relations; the *
 * simulation below stands in for it so the rules can be tested here.   *
 * ------------------------------------------------------------------ */
(function clickResolution() {
  const idx = T.buildIndex([
    '“Business” means the business of designing and selling widgets.',
    '“Business Employee(s)” means each employee of the Business listed on Schedule 2.1(a).',
    '“Business Day” means any day except a Saturday, Sunday or public holiday.',
  ]);
  const para = 'The Buyer shall offer employment to the Business Employees within '
             + 'three Business Days after the Closing, and shall continue to operate the Business.';

  /* What Word reports for a found range [s,e) against a caret at c. */
  const relation = (s, e, c) =>
    c > s && c < e ? 'Contains' :
    c === s ? 'ContainsStart' :
    c === e ? 'ContainsEnd' :
    c < s ? 'After' : 'Before';

  function clickAt(caret) {
    const comparisons = [];
    for (const c of T.candidatesIn(para, idx)) {
      let from = 0, at;
      while ((at = para.indexOf(c.surface, from)) !== -1) {
        comparisons.push({ entry: c.entry, surface: c.surface,
                           relation: relation(at, at + c.surface.length, caret) });
        from = at + 1;
      }
    }
    return T.resolveOverlaps(comparisons);
  }

  const cases = [
    ['inside "Business" of "Business Employees"', para.indexOf('Business Employees') + 3, 'Business Employee(s)'],
    ['inside "Employees"',                        para.indexOf('Employees') + 4,          'Business Employee(s)'],
    ['inside "Business" of "Business Days"',      para.indexOf('Business Days') + 3,      'Business Day'],
    ['inside "Days"',                             para.indexOf('Business Days') + 11,     'Business Day'],
    ['inside standalone "Business"',              para.lastIndexOf('Business') + 3,       'Business'],
    ['inside "Buyer", not a defined term',        para.indexOf('Buyer') + 2,              null],
  ];
  for (const [label, caret, expected] of cases) {
    const r = clickAt(caret);
    const got = r ? r.entry.term : null;
    check('click ' + label, got === expected, `expected ${expected}, got ${got}`);
  }

  // The ambiguous spot must offer the alternative rather than hide it.
  const amb = clickAt(para.indexOf('Business Employees') + 3);
  check('overlapping click offers the other term',
        amb.others.length === 1 && amb.others[0].term === 'Business',
        `got [${amb.others.map(o => o.term).join(', ')}]`);

  // The unambiguous one must not invent alternatives.
  const plain = clickAt(para.lastIndexOf('Business') + 3);
  check('unambiguous click offers nothing extra',
        plain.others.length === 0, `got [${plain.others.map(o => o.term).join(', ')}]`);
})();

/* ------------------------------------------------------------------ *
 * 7. Suspects — the missed-term diagnostic                            *
 * ------------------------------------------------------------------ */
(function conjoinedTermsAreNotNewTerms() {
  const paras = [
    'This Agreement is between ACME INC. (the “Buyer”) and BETA LLC (the “Seller”).',
    'The Buyer and Seller shall cooperate. Seller and Buyer each bear their own costs.',
    'Buyer and Seller agree to the allocation.'
  ];
  const idx = T.buildIndex(paras);
  const sus = T.findSuspects(paras, idx, { minCount: 1 }).map(s => s.phrase);
  check('"Buyer and Seller" is not reported as an unknown phrase',
        sus.indexOf('Buyer and Seller') === -1, `got [${sus.join(', ')}]`);
  check('"Seller and Buyer" is not reported either',
        sus.indexOf('Seller and Buyer') === -1, `got [${sus.join(', ')}]`);
})();

(function compoundsAreStillReported() {
  // Removing a known term must not explain away a genuinely different phrase.
  const paras = [
    '“Administrative Agent” means GLOBAL BANK, N.A.',
    'The Administrative Agent Fee Letter sets out the fees. See the Administrative Agent Fee Letter.'
  ];
  const idx = T.buildIndex(paras);
  const sus = T.findSuspects(paras, idx, { minCount: 1 }).map(s => s.phrase);
  check('"Administrative Agent Fee Letter" is still a suspect',
        sus.indexOf('Administrative Agent Fee Letter') !== -1, `got [${sus.join(', ')}]`);
})();

(function otherTaxesIsNotExplainedAway() {
  // "other" must not count as noise, or a real missed term disappears.
  const paras = [
    '“Taxes” means all present and future taxes imposed by any Governmental Authority.',
    'The Borrower shall pay all Other Taxes when due. Other Taxes are indemnified separately.'
  ];
  const idx = T.buildIndex(paras);
  const sus = T.findSuspects(paras, idx, { minCount: 1 }).map(s => s.phrase);
  check('"Other Taxes" survives as a suspect',
        sus.indexOf('Other Taxes') !== -1, `got [${sus.join(', ')}]`);
})();

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
