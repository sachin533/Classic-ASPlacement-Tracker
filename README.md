# Placement Application Tracker (Classic ASP)

A complete CRUD web application for tracking job and internship applications,
built with **Classic ASP (VBScript)**, **ADO** and **SQL Server Express**, and
styled after the supplied six-screen reference design (dark navy top bar,
white cards on a light blue-grey page, coloured status badges, modal delete
confirmation).

The value of this project is not the legacy framework — every idea here
(routes and HTTP methods, query-string vs form data, server-side validation,
POST/Redirect/GET, status codes, output encoding against XSS, parameterised
SQL against injection, CSRF tokens, sessions, real schema design) carries
over directly to ASP.NET Core, Flask, FastAPI or Node.

---

## 1. What you need

| Requirement | Used on this machine | Notes |
|---|---|---|
| Windows 10/11 | Windows 10 Home Single Language | Home editions have no full IIS — **IIS Express** is used instead. |
| IIS Express + ASP | `C:\Program Files\IIS Express\iisexpress.exe` (includes `asp.dll`) | Runs from an **unelevated** prompt; no admin rights needed. |
| SQL Server Express | instance `.\SQLEXPRESS` (running) | Any SQL Server 2012+ works (the dashboard uses `OFFSET … FETCH`). |
| PowerShell 5+ | built in | Used for the smoke tests and DB verification. |
| (Optional) MS OLE DB Driver | `MSOLEDBSQL` installed | `inc\db.asp` falls back to the built-in `SQLOLEDB` automatically. |

> Full IIS (W3SVC) also works — see *Full IIS deployment* below — but it was
> not testable on this machine (Windows Home, no administrator rights).

---

## 2. One-time database setup

The schema script is **idempotent and non-destructive**: it creates the
database and table only when they do not exist, and never touches rows.

Run **either** of these from the project root:

```powershell
# Option A - PowerShell (verified working on this machine)
Invoke-Sqlcmd -ServerInstance ".\SQLEXPRESS" -InputFile ".\db\schema.sql"
Invoke-Sqlcmd -ServerInstance ".\SQLEXPRESS" -InputFile ".\db\seed.sql"

# Option B - sqlcmd (needs a working ODBC/OLE DB driver on PATH)
sqlcmd -S .\SQLEXPRESS -E -C -b -i db\schema.sql
sqlcmd -S .\SQLEXPRESS -E -C -b -i db\seed.sql
```

* `db\schema.sql` — creates database `PlacementTracker`, table
  `dbo.Applications` with CHECK constraints, and two indexes.
* `db\seed.sql` — inserts the five demonstration rows (Google, Microsoft,
  Amazon, TCS, Infosys) **only when the table is empty**.
* `db\seed_500.sql` — inserts 560 extra realistic sample rows for the search
  and pagination features. Idempotent: sample rows are stamped inside
  2024-01-01, so a re-run inserts nothing.
* `db\verify.sql` — read-only proof of what is actually in the table
  (shape, constraints, integrity violations, newest rows). Runs fine in SSMS;
  its `PRINT` sections are easiest to read there.

Data access uses **Windows integrated security** — no passwords are stored
anywhere. The account running IIS Express (your own user on this machine)
needs read/write access to the database, which the default `.\SQLEXPRESS`
setup already grants to the logged-in user.

---

## 3. Start the site

```powershell
powershell -ExecutionPolicy Bypass -File .\iisexpress\start.ps1
```

Then open **<http://localhost:8080/>**. Stop the server with `Ctrl+C`.

`start.ps1` is safe to re-run: the first execution copies the machine's
default IIS Express configuration into `iisexpress\applicationhost.config`,
adds an application pool (`No Managed Code`, integrated pipeline) and a site
named `PlacementTracker` bound to `localhost:8080`, enables Classic ASP
(`scriptErrorSentToBrowser` is for development only), and launches
`iisexpress.exe`. Binding to the host name `localhost` needs no HTTP.sys URL
reservation, so it works without administrator rights.

Useful parameters: `-Port 9000`, `-DbName PlacementTracker_Test`,
`-GenerateOnly` (config only, no launch).

---

## 4. Run the tests

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\smoke.ps1
```

The suite **never touches the production database**. It rebuilds an isolated
`PlacementTracker_Test` database, starts a second IIS Express instance on
port 8081 with `PT_DB_NAME=PlacementTracker_Test` (read by `inc\db.asp`),
runs the whole suite, then shuts the instance down.

* **Verified result on this machine: 177 passed, 0 failed** (exit code 0).
* Covers: every page loads, 404/405/403 status codes, blocked internal
  folders, CSRF rejection, full CRUD round trip, server-side validation,
  hostile `<script>` / SQL-injection input stored as text and encoded on
  output, Unicode, the empty state, data surviving a server restart,
  idempotent 500-row seeding, column-wide search with LIKE-wildcard
  escaping, numbered pagination, and CRUD combined with search + paging.

`tests\checklist.md` is the manual browser checklist for the same ground.

---

## 5. Project structure

```
placement-tracker\
├── default.asp          Dashboard: search, status filter, pager, table, empty state
├── new.asp              Create: GET form / POST validate + insert + redirect
├── view.asp             Details for one record (404 when unknown)
├── edit.asp             Update: prefilled form / POST validate + update + redirect
├── delete.asp           POST-only delete (GET = 405, CSRF + id re-checked)
├── 404.asp              The site's friendly error page (sets the real status)
├── web.config           default document, custom 404s, security headers, locked folders
├── inc\
│   ├── helpers.asp      UTF-8 switch, H(), ParseId, CSRF, flash, FailPage, icons
│   ├── errorpage.asp    RenderErrorPage() - shared chrome for every error
│   ├── db.asp           connection + parameterised data-access layer
│   ├── validate.asp     CollectApplicationValues() + ValidateApplication()
│   ├── form.asp         RenderApplicationForm() shared by new.asp and edit.asp
│   ├── header.asp       RenderHeader(): doctype, nav, compact search box
│   └── footer.asp       RenderFooter(): footer + shared delete dialog
├── static\
│   ├── styles.css       the reference design (v3)
│   └── app.js           search-clear control + delete dialog (progressive enhancement)
├── db\                  schema.sql, seed.sql, seed_500.sql, verify.sql
├── iisexpress\          start.ps1 + generated applicationhost.config
└── tests\               smoke.ps1 (177 tests) + checklist.md
```

`/inc`, `/db`, `/tests` and `/iisexpress` are unreachable from the browser
(rewrite rules + inner `web.config` files); the smoke tests verify this.

---

## 6. Routes and behaviour

| Method | Page | Purpose | Result |
|---|---|---|---|
| GET | `/default.asp` | List (20 per page, `?q=` search, `?status=`, `?page=`) | 200 |
| GET | `/new.asp` | Blank create form with CSRF token | 200 |
| POST | `/new.asp` | Validate + parameterised INSERT | 302 → `/view.asp?id=` or 200 with errors |
| GET | `/view.asp?id=N` | One record | 200 or 404 |
| GET | `/edit.asp?id=N` | Prefilled edit form | 200 or 404 |
| POST | `/edit.asp?id=N` | Validate + parameterised UPDATE | 302 → details or 200 with errors |
| POST | `/delete.asp` | Delete (`id` + CSRF token in the body) | 302 → dashboard |
| GET | `/delete.asp` | Refuses to delete anything | 405 with `Allow: POST` |

Every POST is rejected with **403** when its CSRF token does not match the
session token. Every unknown or malformed id returns a real **404** through
`404.asp`.

---

## 7. Security model (all verified by the smoke suite)

* **SQL injection** — no SQL is ever built by concatenating user input; every
  command uses positional `?` parameters (`ADODB.Command`).
* **XSS** — every dynamic value is HTML-encoded by `H()` before it is written
  into a page; `<script>` stored through the form is rendered as text.
* **CSRF** — one random token per session, hidden field on every POST form,
  `RequireCsrf` on every state-changing page (403 otherwise).
* **Method safety** — delete is POST-only; a GET (including link prefetch)
  answers 405 and removes nothing.
* **Headers** — `X-Content-Type-Options`, `X-Frame-Options`,
  `Referrer-Policy`, `X-Powered-By` removed.
* **Files** — internal folders are not web-reachable; directory browsing is
  off; the connection string is never echoed to the browser (DB errors log to
  the IIS log via `Response.AppendToLog` and show a generic 500 page).

---

## 8. Configuration

`inc\db.asp` reads two optional **process environment variables**, which is
how the smoke tests point a second server at the isolated test database:

| Variable | Default | Meaning |
|---|---|---|
| `PT_DB_SERVER` | `.\SQLEXPRESS` | SQL Server instance |
| `PT_DB_NAME` | `PlacementTracker` | Database name |

Set them before launching `start.ps1` (its `-DbName` parameter does this for
you). The connection string itself uses Windows integrated security.

---

## 9. Full IIS deployment (not verified on this machine)

On Windows Pro/Enterprise with administrator rights:

1. Enable *Windows Features → Internet Information Services → World Wide Web
   Services → Application Development Features → ASP*.
2. Create a site `PlacementTracker` pointing at this folder, port 8080, app
   pool *No Managed Code*, integrated pipeline.
3. Grant the app-pool identity read access to the folder and map the login
   `IIS APPPOOL\PlacementTracker` to the database with `db_datareader` and
   `db_datawriter` only:
   ```sql
   CREATE LOGIN [IIS APPPOOL\PlacementTracker] FROM WINDOWS;
   USE PlacementTracker;
   CREATE USER [IIS APPPOOL\PlacementTracker] FOR LOGIN [IIS APPPOOL\PlacementTracker];
   ALTER ROLE db_datareader ADD MEMBER [IIS APPPOOL\PlacementTracker];
   ALTER ROLE db_datawriter ADD MEMBER [IIS APPPOOL\PlacementTracker];
   ```
4. `web.config` already provides the default document, custom 404s, the
   locked-down folders and the security headers.

---

## 10. Troubleshooting

| Symptom | Fix |
|---|---|
| `http://localhost:8080/` refuses the connection | Another process holds port 8080 (check with `netstat -ano \| findstr :8080`), or the IIS Express window was closed — re-run `start.ps1`, optionally with `-Port 9000`. |
| “We could not complete that request” (generic 500) | SQL Server is stopped, or the database does not exist yet — run `db\schema.sql`. Details are in the IIS Express console/log (`DBERROR`). |
| `Provider or data source not found` in the log | `MSOLEDBSQL` is not installed — the built-in `SQLOLEDB` fallback should engage automatically; if it does not, install the *Microsoft OLE DB Driver for SQL Server*. |
| Pages show stale styling after a CSS edit | The stylesheet is versioned (`styles.css?v=3`) — bump the version in `inc\header.asp`. |
| `smoke.ps1` exits with code 2 | The test server on port 8081 could not start — another process holds the port; re-run or pass `-Port 8082`. |
