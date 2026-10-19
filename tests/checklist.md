# Manual test checklist — Placement Application Tracker

Run against a development site (`.\iisexpress\start.ps1` → http://localhost:8080/).
Tick every box. The automated equivalent is `tests\smoke.ps1` (177 tests) —
this checklist exists so a human can confirm the *feel* of each screen
against the reference design as well as the behaviour.

Legend: [ ] not run · [x] passed · [!] failed

---

## 1. Pages load (reference design)

- [ ] `/` — dark navy top bar, "Placement Tracker" brand, Home underlined,
      compact white search box on the right, light blue-grey page background.
- [ ] Table shows `# | Company | Role | Status | Applied On | Actions` with
      coloured pill badges: Applied = blue, Interview = purple, Offer =
      green, Rejected = red, Wishlist = grey.
- [ ] "Applications" heading with subtitle; blue **+ Add Application** button
      top-right.
- [ ] `/new.asp` — "← Back to Applications", two-column form grid, red
      asterisks on required fields, blue **Save Application** + grey
      **Cancel** buttons.
- [ ] Row "View" opens the details page; "Edit" opens the edit form.
- [ ] `/static/styles.css` and `/static/app.js` return 200.
- [ ] Narrow the window (< 760 px): form becomes one column, search box drops
      under the nav, table scrolls horizontally.

## 2. Create

- [ ] Submit the empty form → red summary + per-field messages, **nothing**
      is saved, and every value you *did* type is still in the form.
- [ ] Company `A` (one character) → rejected.
- [ ] Status `Applied` with no date → date error; status `Wishlist` with no
      date → accepted.
- [ ] Job URL `javascript:alert(1)` → rejected; `https://…` → accepted.
- [ ] Valid row → redirected to its details page, green flash
      "Application created successfully.", new row visible on the dashboard.

## 3. Read

- [ ] Details page lists Company, Role, Status badge, Applied On, Job URL
      (blue link), Notes, Created At (UTC).
- [ ] `/view.asp?id=999999` → 404 status **and** the friendly 404 card.
- [ ] `/view.asp?id=abc`, `/view.asp?id=-5`, `/view.asp?id=1 OR 1=1` → 404.
- [ ] Unknown path `/nonsense` → 404 page (not an IIS error).

## 4. Edit

- [ ] Edit form arrives prefilled, including the date in the date picker.
- [ ] Change Applied → Interview, save → redirected to details, badge is
      purple, flash "Application updated successfully."
- [ ] POST to `/edit.asp?id=999999` (e.g. stale form) → 404, nothing changed.

## 5. Delete

- [ ] Dashboard "Delete" opens the **modal**: title "Delete Application",
      red warning circle, "Yes, Delete" (red) and "Cancel" (grey).
- [ ] Cancel / × / clicking outside / pressing Esc closes it, nothing deleted.
- [ ] "Yes, Delete" → back to dashboard, green flash, row gone.
- [ ] Details-page Delete button opens the same modal.
- [ ] `GET /delete.asp` in the address bar → **405**, row count unchanged.
- [ ] `POST /delete.asp` without a CSRF token (curl/PowerShell) → 403.
- [ ] Delete a record in a second tab, then delete it in the first → 404
      message, no crash.

## 6. Search, filter, pagination

- [ ] Search `Google` → only matching rows; summary shows the count; chip
      "Search: Google" with "Clear all".
- [ ] Search `%` matches nothing (wildcard escaped, not "everything").
- [ ] Status filter `Offer` → only offers; search + filter combine.
- [ ] 560-row seed: pager shows `1 … 27 28 29`, Previous/Next states correct;
      the `#` column continues across pages (21–40 on page 2).
- [ ] New search resets to page 1; page links keep the query.
- [ ] Header search box preserves the query.

## 7. Empty state

- [ ] Delete every row (or use the test DB) → briefcase icon in a pale blue
      circle, "No applications yet", support text, **+ Add Application**
      button, navigation still present, no table rendered.

## 8. Security and robustness

- [ ] Enter company `<script>alert(1)</script>` → stored as text, no alert
      on dashboard or details.
- [ ] Search box `<script>alert(1)</script>` → encoded, never executed.
- [ ] `/inc/db.asp`, `/db/schema.sql`, `/tests/smoke.ps1` → 404 (blocked).
- [ ] Response headers: no `X-Powered-By`; `X-Content-Type-Options: nosniff`
      present.
- [ ] Unicode company name (e.g. `Ünïcodé Corp`) saves and displays intact.
- [ ] Restart the site (`Ctrl+C`, `start.ps1`) → all data still present.

## 9. Database verification

- [ ] `db\verify.sql` integrity sections return **zero** rows:
      non-Wishlist without date, non-http(s) URLs, bad lengths, bad statuses.
- [ ] `seed.sql` / `seed_500.sql` can both be re-run without duplicating rows.
