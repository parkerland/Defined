# Defined

A Microsoft Word task pane add-in that shows the definition of a defined term
when you click it in the document.

Built for reviewing long legal and commercial documents. Plain HTML, CSS and
JavaScript — no build tools, no frameworks, no npm. Everything runs inside the
task pane; **no document content is sent anywhere.**

---

## What each file does

| File | What it is |
|---|---|
| `terms.js` | **The parser.** Takes paragraphs of text, returns defined terms. Knows nothing about Word — which is why you can test it in a browser. |
| `app.js` | Everything that talks to Word: scanning, cursor tracking, drawing the pane. |
| `index.html` | The task pane itself. |
| `taskpane.css` | Styling for the pane. |
| `test.html` | **Pattern test bench.** Open in any browser, paste in a real agreement, see what gets found. No Word needed. |
| `patterns.test.js` | Regression tests for the parser. Run `node patterns.test.js`. |
| `manifest.xml` | Tells Word the add-in exists and where it lives. This is the file you install. |
| `commands.html` | Required by the manifest. Does nothing yet. |
| `assets/` | Icons. |
| `.nojekyll` | Stops GitHub Pages pre-processing the files. Optional but harmless. |

---

## Step 1 — Try the parser first (5 minutes, no Word)

Double-click **`test.html`**. It opens in your browser with sample
credit-agreement text already loaded.

You'll see three sections:

- **Terms found** — what the patterns matched, and which pattern matched.
- **Quoted text that was *not* indexed** — ordinary quotations belong here.
  **If a real defined term shows up in this list, that's a pattern to improve.**
- **Click simulator** — checks the matching logic used for a double-click.

Now go to [SEC EDGAR](https://www.sec.gov/edgar/search/), open a credit
agreement, select all, copy, and paste it into the box. Press **Parse**.

Do this before touching Word. Tuning the patterns is most of the work, and here
the loop is "edit `terms.js`, refresh the page" instead of "re-upload, restart
Word, clear cache".

---

### Finding terms the parser missed

Scroll to **"Capitalised phrases the index can't explain"**. Any repeated Title
Case phrase that isn't a known term (or part of one, or its plural) is listed,
split into two groups:

- **A. Looks like a definition the parser missed** — the phrase is followed by
  *means* / *shall mean* / *has the meaning*, but it never got indexed. **These
  are parser bugs.** Use the copy button and report them.
- **B. Used but never defined** — no defining sentence anywhere. Mostly party
  names and statutes, but it doubles as an undefined-terms check.

### Before changing the patterns

Run `node patterns.test.js` in the project folder. 33 checks, under a second,
no setup.

The patterns are deliberately loose, and loosening one to catch a missed term
is an easy way to start matching things that aren't definitions. The test file
has as many *negative* cases (quoted prose that must **not** be indexed) as
positive ones. Run it before and after any change to `terms.js`.

---

## Step 2 — Put the files on GitHub Pages

Word requires add-ins to be served over HTTPS. GitHub Pages does that free.

1. Go to <https://github.com/new>.
2. Repository name: **`Defined`** (the capital D matters — it becomes part of
   the web address, and those are case-sensitive).
3. Choose **Public**. Free GitHub Pages requires it. The repository holds only
   this code — never any document you review.
4. Click **Create repository**.
5. On the next page click **uploading an existing file**.
6. Open `C:\Users\alex\Claude_Code\Defined` in File Explorer, select everything
   including the `assets` folder, and drag it into the browser window.
7. Click **Commit changes**.
8. Go to **Settings → Pages**. Under *Build and deployment*, set
   **Source: Deploy from a branch**, **Branch: `main`**, folder **`/ (root)`**.
   Click **Save**.
9. Wait 1–2 minutes, then open <https://parkerland.github.io/Defined/test.html>.
   If the test bench loads, your hosting works.

> If you see a 404, give it another minute — the first publish is the slowest.
> Check **Settings → Pages** again; it shows a green "Your site is live at…"
> banner when it's ready.

---

## Step 3 — Install in Word on the web

Test here first. It's faster and has proper developer tools.

1. Upload a test document to OneDrive and open it in **Word for the web**.
2. On the **Home** tab, click **Add-ins**.
3. Click **More Add-ins** (then the **My Add-ins** tab if you're not already on it).
4. Click **Upload My Add-in**, top right.
5. **Browse** to your local `manifest.xml` and click **Upload**.
6. The pane opens and starts scanning. The header shows e.g. "412 terms indexed".

Click a defined term in the document — the definition appears in the pane.

> You're uploading the manifest from your own disk, not from GitHub. The
> manifest just tells Word to load the pane *from* GitHub.

---

## Step 4 — Install in Word on Windows desktop

Desktop Word only sideloads add-ins from a **shared network folder**, so there's
one extra setup step. You do this once.

**a. Make a shared folder**

1. Create `C:\WordAddins`.
2. Copy `manifest.xml` into it.
3. Right-click the folder → **Properties** → **Sharing** tab → **Share…**
4. Add your own user account, give it **Read** permission, click **Share**,
   then **Done** and **Close**.

**b. Tell Word to trust it**

1. Open Word. Go to **File → Options → Trust Center → Trust Center Settings…**
2. Choose **Trusted Add-in Catalogs**.
3. In **Catalog Url**, type `\\localhost\WordAddins` and click **Add catalog**.
4. Tick **Show in Menu** next to the entry you just added.
5. **OK**, **OK**, then **close Word completely and reopen it**.

**c. Insert the add-in**

1. **Insert** tab → **Add-ins** (or **My Add-ins**).
2. Click the **SHARED FOLDER** tab.
3. Select **Defined** → **Add**.

After this, the add-in appears on the **Home** tab as a **Defined terms** button.

---

## Using it

- **Click a term** in the document. If a bare click doesn't catch it,
  **double-click** the word — that's the most reliable gesture.
- **Follow cursor** — untick this to browse the term list without the cursor
  changing what's shown.
- **Search box** — filter the full list of defined terms.
- **Refresh** — re-scan after editing the document. The add-in does *not*
  re-scan automatically.
- **Diagnostics** (at the bottom) — scan timings and any errors. Expand this
  first if something looks wrong.

A term is matched in every form it might be written: plurals
("Subsidiaries" finds "Subsidiary"), singulars ("Loan Document" finds "Loan
Documents"), possessives ("Borrower's"), and the bracketed notation drafters
use ("Business Employees" finds "Business Employee(s)").

Each definition is tagged with how it was found:

| Tag | Meaning |
|---|---|
| `means` | `"X" means …` / `shall mean` — the strongest match |
| `cross-reference` | `"X" has the meaning given in Clause 5.2` |
| `inline` | `(the "X")` — defined in passing |
| `inline (uncertain)` | An inline match the parser is less sure about |

---

## When you change the code

GitHub Pages and Word both cache aggressively. **This is the single most likely
thing to waste your time** — you change a file, nothing happens, and you assume
the change was wrong.

- **Changed `app.js`, `terms.js` or `taskpane.css`?** Open `index.html` and bump
  the version numbers: `app.js?v=1` becomes `app.js?v=2`. Do it for every
  changed file. That forces a fresh download.
- **Changed `manifest.xml`?** Bump `<Version>1.0.0.0</Version>` to `1.0.0.1`,
  then remove and re-add the add-in.
- **Still stale on desktop?** Close Word, then delete everything inside
  `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` and reopen Word. Paste that path
  into the File Explorer address bar.
- **Still stale on the web?** Hard-refresh the browser with `Ctrl+Shift+R`.

To see JavaScript errors: in Word on the web, press `F12` for the browser
console. On desktop, right-click inside the task pane and choose **Inspect**.

---

## Privacy

- Document text is read into the task pane, parsed there, and discarded when you
  close the pane. It's never written anywhere or transmitted.
- There are no `fetch` or `XMLHttpRequest` calls in `app.js` or `terms.js`. You
  can verify that yourself: search those files for `fetch` and `http`.
- Two files *are* loaded over the network, both at startup, neither carrying
  document content: the pane's own files from GitHub Pages, and the standard
  `office.js` library from Microsoft's CDN (every Office add-in loads this).
- GitHub's servers see requests for your HTML and JavaScript files, with your IP
  address, the same as any website. They never see the document.

---

## Not in version 1

Deliberately left out, to be added once the basics are solid:

- Clickable nested terms inside a definition
- Go-to-definition (jump to where the term is defined)
- Following `has the meaning given in Clause X` through to the actual text
- Flagging duplicate, unused, or undefined-but-capitalised terms

Groundwork for several of these is already in place: each entry records its
paragraph number, any clause pointer it refers to, and any competing definitions
found elsewhere in the document.
