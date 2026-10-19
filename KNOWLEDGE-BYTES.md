# Knowledge Bytes — Placement Application Tracker (Classic ASP)

Small, self-contained explanations of the real code in this project, ordered
so that understanding builds up: shared foundation → database layer →
dashboard → CRUD pages. Every snippet is copied from the file named below it.

Reader level: you know basic programming but are new to Classic ASP/VBScript,
ADO and this codebase. Terms are defined the first time they appear.

---

# Module 1 — Shared foundation (`inc\`)

## Byte 1: Includes aur shared page state

**Builds on:** None — starting point

**In plain terms:**
Classic ASP me har page ek independent script hai. Lekin har page ko same
navigation, same encoding aur same helpers chahiye — isliye `<!--#include-->`
directive ek file ka *text* dusre file me copy-paste jaisa insert kar deta
hai. `helpers.asp` sabse pehle include hota hai kyunki yeh poori response ko
UTF-8 me switch karta hai aur ek *page-scope* variable banata hai jise header
padhta hai.

**The code:** `inc\helpers.asp` (top of the file)

```vbs
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"

' -- Shared page state (read by inc\header.asp and inc\errorpage.asp) -------
Dim gTitle, gActiveNav, gShowSearch, gSearchValue, gFilterStatus
Dim gErrCode, gErrTitle, gErrMsg

gTitle       = ""
gActiveNav   = ""
gShowSearch  = False
```

**What's happening:**
Include = textual splice. Isi liye ek file ko ek page me *do baar* include
nahi kar sakte (duplicate `Sub`/`Const` error). `CodePage 65001` matlab
server UTF-8 me likhe — tabhi `Ünïcodé Corp` jaisi values sahi store aur
display hoti hain. `g*` variables *shared page state* hain: har page apna
`gTitle` set karta hai, aur `RenderHeader()` (header.asp) wahi padh kar
`<title>` aur active nav link banata hai.

**Why it matters:**
Layout ek hi jagah se aata hai — navigation badalna sirf `header.asp` me
ek line hai. UTF-8 na ho to non-English names database me theek se store
nahi hote (mojibake).

---

## Byte 2: `H()` — har output ka HTML encoder

**Builds on:** Byte 1

**In plain terms:**
`H()` har dynamic value ko HTML me likhne se *pehle* encode karta hai:
`<` becomes `&lt;`, `"` becomes `&quot;`. Iske bina koi user apne notes me
`<script>alert(1)</script>` likh kar doosre users ke browser me script chala
sakta tha — yahi attack hai **XSS** (Cross-Site Scripting).

**The code:** `inc\helpers.asp`

```vbs
' Encode anything before it is written into HTML.
Function H(v)
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    H = Server.HTMLEncode(CStr(v))
End Function
```

**What's happening:**
`Null`/`Empty` ko `""` bana deta hai (warna `Server.HTMLEncode(Null)` error
deta), aur phir `Server.HTMLEncode` kaam karta hai. Dashboard ke table me
company naam ho ya details page ka notes — sab `H()` se guzarte hain.

**Why it matters:**
Smoke test me prove hua hai: `<script>` company naam *store* to text ki tarah
hoti hai, par page par dikhta hai `&lt;script&gt;` — browser use text dikhata
hai, code nahi chalata. `H()` kisi bhi page par miss ho gaya = XSS hole.

---

## Byte 3: `ParseId` — link ke id ko safely padhna

**Builds on:** Byte 2

**In plain terms:**
Browser URL me user kuch bhi type kar sakta hai: `?id=abc`, `?id=-5`,
`?id=1 OR 1=1`. Har page ek *clean non-negative integer* chahta hai. Strict
regex se check karke hi hum `CLng()` (string → integer) lagate hain, warna
VBScript crash ya absurd lookups ho sakte hain.

**The code:** `inc\helpers.asp`

```vbs
Function ParseId(raw)
    ParseId = -1
    Dim s
    s = Trim(raw & "")
    If s = "" Then Exit Function
    If Not TestPattern("^[0-9]{1,9}$", s) Then Exit Function
    Dim n
    n = CLng(s)
    If n < 0 Then Exit Function
    ParseId = n
End Function
```

**What's happening:**
Return value `-1` ka matlab "invalid". `IsNumeric()` kafi nahi hai — woh
`"1E5"` ya `"-1"` ko bhi numeric maanta hai. `^[0-9]{1,9}$` sirf 1 se 9
digits allow karta hai (integer overflow se bachne ke liye). Note: **id 0
valid hai** — comment me explain hai ki `DBCC CHECKIDENT ... RESEED, 0`
genuinely row id 0 bana sakta hai.

**Why it matters:**
`view.asp?id=abc` par page 500 ki jagah graceful 404 deta hai — yahi
"unknown or non-numeric ID → proper 404" requirement hai. Injection strings
(`1 OR 1=1`) bhi yahin filter ho jati hain.

---

## Byte 4: CSRF token — POST ko apne-browser se prove karna

**Builds on:** Byte 1

**In plain terms:**
**CSRF** (Cross-Site Request Forgery) attack me attacker aapke browser se
aapki session ke saath galat POST bhejwata hai (jaise ek page jo aapke
tracker par delete trigger kare). Defence: har session ka ek random token,
jo har POST form me hidden field me jata hai, aur server har POST par
match karta hai — mismatch = **403**.

**The code:** `inc\helpers.asp`

```vbs
Function CsrfToken()
    Dim t
    t = Session("csrf")
    If IsNull(t) Or IsEmpty(t) Then t = ""
    If t = "" Then
        t = NewToken()
        Session("csrf") = t
    End If
    CsrfToken = t
End Function

Sub RequireCsrf()
    Dim posted
    posted = Request.Form("csrf") & ""
    If StrComp(posted, CsrfToken(), vbBinaryCompare) <> 0 Then
        FailPage 403, "Request blocked", _
                 "Your session token was missing or is no longer valid. ..."
    End If
End Sub
```

**What's happening:**
`Session` IIS ka per-visitor memory box hai (yahan 30 minute). Token pehli
bar zaroorat par `NewToken()` (GUID generator, with fallback) se banta hai
aur session me rehta hai. Form me `CsrfInput()` hidden input chhapta hai;
POST par `RequireCsrf` form value ko session value se **binary** compare
karta hai (timing-safe nahi, par sufficient).

**Why it matters:**
Smoke tests verify karte hain: token ke bina ya galat token ke saath POST →
403 **aur database me zero rows badli**. Delete sirf wahi browser kar sakta
hai jiske paas page khula tha.

---

## Byte 5: `FailPage` — asli HTTP status ke saath friendly page

**Builds on:** Byte 1, Byte 3

**In plain terms:**
"404 Not Found" sirf ek page nahi — ek *HTTP status code* hai jo browser,
search engines aur test scripts padhte hain. `FailPage` pehle response ka
buffer clear karta hai, phir `Response.Status` set karta hai, phir same
chrome (header/footer) ke saath error card render karta hai.

**The code:** `inc\helpers.asp`

```vbs
Sub FailPage(code, title, message)
    Response.Clear
    On Error Resume Next
    Response.TrySkipIisCustomErrors = True
    On Error GoTo 0

    Select Case CLng(code)
        Case 403 : Response.Status = "403 Forbidden"
        Case 404 : Response.Status = "404 Not Found"
        Case 405 : Response.Status = "405 Method Not Allowed"
                   Response.AddHeader "Allow", "POST"
        Case 500 : Response.Status = "500 Internal Server Error"
    End Select
    ...
    RenderErrorPage CLng(code), title, message
    Response.End
End Sub
```

**What's happening:**
`TrySkipIisCustomErrors` IIS ko bolta hai "humara apna body mat badlo" —
warna `web.config` ka `<httpErrors>` response body ko apne default error
page se replace kar deta. `Response.End` ke baad page ka koi code nahi
chalta. `errorpage.asp` ke header me explain hai ki `Server.Execute` yahan
kyun nahi use kiya (query string nahi le sakta, aur naya page scope kholta
hai — isliye code/title/message arguments me pass hote hain).

**Why it matters:**
Smoke suite ke status-code tests (`view.asp?id=999999 → 404`,
`GET /delete.asp → 405`) isi helper par khade hain. Status galat ho to
"page dikh raha hai" par tools ko lagega sab theek hai.

---

## Byte 6: Flash message — redirect ke baad ek message

**Builds on:** Byte 1 (Session)

**In plain terms:**
Save karne ke baad hum *redirect* karte hain (POST/Redirect/GET pattern).
Redirect response me koi body nahi hoti — to "saved successfully" kahan
dikhayein? Trick: message ko `Session` me rakho, agle page (details page)
par dikhao aur turant mita do.

**The code:** `inc\helpers.asp` + `inc\header.asp`

```vbs
Sub SetFlash(kind, message)
    Session("flash_kind")    = kind & ""
    Session("flash_message") = message & ""
End Sub

Function RenderFlash()
    ...
    kind    = Session("flash_kind")
    message = Session("flash_message")
    Session("flash_kind")    = ""      ' one-shot: clear after reading
    Session("flash_message") = ""
    ...
```

**What's happening:**
`new.asp` save hote hi `SetFlash "success", "Application created
successfully."` call karta hai aur `view.asp?id=N` par redirect hota hai.
`view.asp` ka `RenderHeader()` andar `RenderFlash()` bulata hai — message
dikh kar turant session se mit jata hai (isliye refresh par dobara nahi
aata).

**Why it matters:**
Bina flash ke user ko pata hi nahi ki save hua ya nahi; POST directly
render karte to browser-refresh par form dobara submit ho sakta tha
(double-insert). Yehi PRG pattern ka core benefit hai.

---

### PUTTING IT TOGETHER (Module 1)

Har page ka skeleton same hai: `helpers.asp` pehle (UTF-8 + shared state +
sare helpers), phir `db.asp`, `validate.asp`, `form.asp`. Header `gTitle`
padh kar nav/title banata hai; har user value `H()` se encode hoti hai;
har id `ParseId` se; har POST `RequireCsrf` se; har error `FailPage` se real
status code ke saath; har successful save `SetFlash` + redirect. Yeh ek
chhota "framework" hai jispar ab CRUD pages khade honge.

---

# Module 2 — Database layer (`inc\db.asp`)

## Byte 7: `OpenDb` — connection string, provider fallback, env override

**Builds on:** Module 1

**In plain terms:**
Database se baat karne ke liye ek *connection* chahiye. Yahan connection
string me koi password nahi — *Windows integrated security* matlab IIS
Express jo user chalata hai (aap), wahi SQL Server par apni Windows pehchan
se login karta hai. Do providers hain: naya `MSOLEDBSQL` (preferred) aur
purana built-in `SQLOLEDB` (fallback).

**The code:** `inc\db.asp`

```vbs
Dim PT_SERVER, PT_DATABASE
PT_SERVER   = ".\SQLEXPRESS"
PT_DATABASE = "PlacementTracker"
...
ptEnvServer   = Trim(PtEnvVar("PT_DB_SERVER"))
ptEnvDatabase = Trim(PtEnvVar("PT_DB_NAME"))
If ptEnvServer <> ""   Then PT_SERVER   = ptEnvServer
If ptEnvDatabase <> "" Then PT_DATABASE = ptEnvDatabase
```

**What's happening:**
Process environment variables (`PT_DB_SERVER`, `PT_DB_NAME`) defaults ko
override kar sakti hain — yahi woh hook hai jisse `tests\smoke.ps1` dusra
server instance `PlacementTracker_Test` database par chala kar production DB
ko bilkul chhota nahi. `OpenDb()` pehle primary provider try karta hai,
fail ho to fallback list traverse karta hai, aur kuch na chale to friendly
500 (`DbFail`) — connection string kabhi browser me nahi jata.

**Why it matters:**
"Data survives an IIS restart" isliye possible hai ki data aapke browser me
nahi, SQL Server me hai. Env-override ne tests ko production se isolate
kiya bina code badle.

---

## Byte 8: Positional `?` parameters — injection ka ilaaj

**Builds on:** Byte 7

**In plain terms:**
SQL *injection* tab hota hai jab user ka text SQL statement me directly
jodha jata ho (`"... WHERE company = '" & user & "'"`). Isko rokne ka sahi
tareeka *parameterised query* hai: SQL me holes (`?`) likho, values alag se
bhejo. Driver values ko escape karke bhejta hai — user ka `' OR 1=1 --`
sirf ek string ban kar rehta hai.

**The code:** `inc\db.asp` (used by insert, update and delete)

```vbs
Sub AppendApplicationParameters(cmd, v)
    cmd.Parameters.Append cmd.CreateParameter("company", adVarWChar, 1, 100, v("company"))
    cmd.Parameters.Append cmd.CreateParameter("role",    adVarWChar, 1, 100, v("role"))
    cmd.Parameters.Append cmd.CreateParameter("status",  adVarWChar, 1, 20,  v("status"))

    Dim pDate
    Set pDate = cmd.CreateParameter("applied_on", adDBTimeStamp, 1, 8)
    ' Leaving a parameter's Value unassigned sends an EMPTY value, not NULL...
    If v("applied_on_iso") <> "" Then
        pDate.Value = CDate(v("applied_on_iso"))
    Else
        pDate.Value = Null
    End If
    cmd.Parameters.Append pDate
    ...
```

**What's happening:**
Do critical details: (1) placeholder *positional* (`?`) hain, named nahi —
file ke comment me explain hai ki MSOLEDBSQL named parameters (`@id`) ko
verbatim pass kar deta hai aur SQL Server reject kar deta hai; isliye
parameters *statement order me* append hone chahiye. (2) Khaali optional
field ko `""` nahi bhejte — `Null` set karta hai, kyunki `''` ek valid DATE
nahi hai aur CHECK constraint fail hota.

**Why it matters:**
Smoke suite me prove hua: `' OR 1=1 --` company naam ki tarah *store* hua,
table intact rahi. Agar kisi ek query me bhi concatenation ho jaye, poora
defence gir jata hai — isi liye `db.asp` ka rule hai "no request value is
ever concatenated into SQL".

---

## Byte 9: `EscapeLike` — search ke wildcards ko literal karna

**Builds on:** Byte 8

**In plain terms:**
Search `LIKE '%term%'` use karta hai, jahan `%` aur `_` special (wildcard)
hain. Agar user search box me literal `%` type kare to usse *sab kuch*
match nahi karna chahiye — use `%` naam wala text hi dhoondhna chahiye.
`EscapeLike` user input ke in special characters ko bracket-escape karta
hai.

**The code:** `inc\db.asp`

```vbs
Function EscapeLike(ByVal s)
    s = Replace(s, "[", "[[]")
    s = Replace(s, "%", "[%]")
    s = Replace(s, "_", "[_]")
    EscapeLike = s
End Function
```

**What's happening:**
`[%]` SQL Server ka character-class hai jiska matlab hai "ek literal percent
sign". Order matter karta hai: `[` pehle escape hota hai warna hum khud
naya bracket insert kar dete. `ByVal` keyword critical hai — comment me
bataya gaya bug: ek purane `ByRef` version me caller ka search term har call
par badalta tha aur paancho LIKE parameters alag-alag pattern dekhne lagte
the.

**Why it matters:**
Tests verify: search `%` ya `_` ya `[` → 200 status aur **zero** results
(matlab literal match, wildcard nahi). Yeh "search is injection safe"
section ka hissa hai.

---

## Byte 10: `FetchApplicationsPage` — COUNT + ek page, ek connection par

**Builds on:** Byte 8, Byte 9

**In plain terms:**
Dashboard ko har page par sirf 20 rows chahiye, par "of 132" wala total
bhi dikhana hai. SQL Server 2012+ ka `OFFSET … FETCH NEXT` pagination
deta hai. Fazaal: page number user se aata hai (`?page=99999`) — use
*clamp* karna zaroori hai warna khaali page ya error aata hai.

**The code:** `inc\db.asp`

```vbs
    ' --- 2. clamp the requested page so ?page=99999 can never 500 or 404 ---
    pages = Int((total + pageSize - 1) / pageSize)
    If pages < 1 Then pages = 1
    page = CLng(requestedPage + 0)
    If page < 1 Then page = 1
    If page > pages Then page = pages

    offsetRows = (page - 1) * pageSize
    ...
    cmd.CommandText = "SELECT id, company, ... FROM dbo.Applications" & _
                      ApplicationWhere(term, statusFilter) & _
                      " ORDER BY id DESC " & _
                      "OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
```

**What's happening:**
COUNT query aur page query *same* `ApplicationWhere()` fragment use karte
hain — isliye total hamesha poori filtered set batata hai, sirf screen ki
rows nahi. `ORDER BY id DESC` me unique id tie-breaker hai, warna do rows
ka order page ke beech me palat sakta tha. Offset aur page size *bhi*
parameters hain (arithmetic se bane integers, raw text kabhi nahi).
Result ek `Scripting.Dictionary` hai: `rows`, `total`, `page`, `pages`,
`firstNo`, `lastNo`.

**Why it matters:**
560 rows par bhi dashboard 20 rows render karta hai aur page 29 tak ka pager
sahi kaam karta hai. Page clamping ke bina `?page=-3` ya `?page=99999`
hand-edled URL 500 de sakta tha — test suite in dono ko check karta hai.

---

## Byte 11: `OUTPUT INSERTED.id` — naye record ki ID wapas lena

**Builds on:** Byte 8

**In plain terms:**
Insert ke baad hum user ko `view.asp?id=N` par bhejte hain — matlab nayi
row ki ID chahiye. Classic tareeka `SCOPE_IDENTITY()` hai, par yahan woh
kaam nahi karta (dekhiye comment). Isliye SQL Server ke `OUTPUT INSERTED.id`
clause ka use hai jo INSERT ke andar hi nayi ID result-set me lauta deta
hai.

**The code:** `inc\db.asp`

```vbs
    cmd.CommandText = "INSERT INTO dbo.Applications " & _
                      "(company, role, status, applied_on, job_url, notes) " & _
                      "OUTPUT INSERTED.id " & _
                      "VALUES (?, ?, ?, ?, ?, ?)"
    ...
    ' OUTPUT INSERTED.id returns the new key as a resultset from the INSERT
    ' itself.  SCOPE_IDENTITY() is deliberately NOT used: ADO sends a
    ' parameterised command through sp_prepexec/sp_execute, which is a nested
    ' scope, so a following SCOPE_IDENTITY() call reads NULL.
```

**What's happening:**
`cmd.Execute` yahan ek recordset return karta hai jiska pehla column nayi ID
hai (`CLng(rs(0).Value)`). `@@IDENTITY` se behtar hai kyunki woh kisi aur
trigger ki identity bhi la sakta hai; `OUTPUT` sirf isi statement ki
`INSERTED` row deta hai — concurrent inserts se bhi safe.

**Why it matters:**
Yehi "POST/Redirect/GET to the new record" requirement ko possible banata
hai. `SCOPE_IDENTITY()` par gaye ho to har successful insert ke baad
`Nothing` par crash ya redirect-to-blank milta — smoke suite ka
"Redirects to the new details page" test isi par khada hai.

---

### PUTTING IT TOGETHER (Module 2)

`db.asp` ek chhota data-access layer hai: `OpenDb()` connection banata hai
(env override ke saath), `AppendApplicationParameters` chhe fields ko
order me bind karta hai (khaali optional fields ko NULL banate hue), aur
upar ke functions — `FetchApplicationsPage`, `GetApplication`, `InsertApplication`,
`UpdateApplication`, `DeleteApplication` — pages ko *sirf dictionaries aur
row counts* dete hain. Pages ko SQL nahi dikhti; isi wajah se pagination,
search aur CRUD ek hi jagah fix ho jate hain aur har page chhota rehta hai.

---

# Module 3 — Dashboard (`default.asp` + `static\`)

## Byte 12: Dashboard ka data flow — query string se HTML tak

**Builds on:** Module 1, Module 2

**In plain terms:**
Dashboard teen query-string parameters padhta hai: `?q=` (search), `?status=`
(filter), `?page=` (page number). Inhe normalise karke `FetchApplicationsPage`
ko deti hai, jo dictionary lauti hai — rows, total, pages. Phir ek `For`
loop har row ko table me chhapta hai.

**The code:** `default.asp`

```vbs
searchRaw = Trim(Request.QueryString("q") & "")
If Len(searchRaw) > 100 Then searchRaw = Left(searchRaw, 100)

statusRaw    = Trim(Request.QueryString("status") & "")
statusFilter = CanonicalStatus(statusRaw)   ' "" when blank or not one of the five

requestedPage = ParsePageNumber(Request.QueryString("page"))
...
Set result = FetchApplicationsPage(searchRaw, statusFilter, requestedPage, PT_PAGE_SIZE)
Set apps   = result("rows")
total      = result("total")
```

**What's happening:**
`Request.QueryString` URL ke parameters ka door (dictionary) hai — yeh
browser par bhejta hai jab user search kare ya page badle. `Request.Form`
uska counterpart hai (form submit par milta hai — Byte 17 me). Har input
pehle bounded/truncated hota hai, `CanonicalStatus` sirf paanch allowed
statuses me se ek return karta hai (unknown value = ""), aur page number
`ParsePageNumber` se safe hota hai. Loop `apps.Item(CStr(i))` se row
nikal kar `<td>` me `H()` ke saath chhapta hai; status ke liye
`StatusBadge()` (Byte 13) aur date ke liye `row("applied_on")` jo
`db.asp` me hi ISO string ban chuki hai.

**Why it matters:**
Yehi "trace dashboard data from database to HTML" wala flow hai. Input
boundary par sanitise hone se aage ka code sirf "already safe" values par
depend karta hai.

---

## Byte 13: Status badges — reference design ka palette

**Builds on:** Byte 12

**In plain terms:**
Paanch statuses ko coloured pill me dikhana hai (Applied = neela, Interview
= baingani, Offer = hara, Rejected = laal, Wishlist = gri). `StatusBadge`
function HTML string banata hai; CSS classes `styles.css` me define hain.

**The code:** `inc\helpers.asp`

```vbs
Function StatusBadge(status)
    Dim key, cls
    key = LCase(Trim(status & ""))
    Select Case key
        Case "wishlist"  : cls = "badge badge-wishlist"
        Case "applied"   : cls = "badge badge-applied"
        Case "interview" : cls = "badge badge-interview"
        Case "offer"     : cls = "badge badge-offer"
        Case "rejected"  : cls = "badge badge-rejected"
        Case Else        : cls = "badge badge-unknown"
    End Select
    StatusBadge = "<span class=""" & cls & """>" & H(status) & "</span>"
End Function
```

**What's happening:**
`LCase` comparison taaki case-insensitive ho; `Class Else` safety net hai.
Color CSS ka kaam hai (`.badge-applied { background: var(--badge-applied); }`
etc.), markup me sirf class name aata hai — isi liye palette badalna ek CSS
line hai. Badge ka text bhi `H()` se guzarta hai.

**Why it matters:**
Reference design ki sabse pehchan yehi coloured badges hain. CSS aur HTML
alag-alag hone se designer ko VBScript chhune ki zaroorat nahi.

---

## Byte 14: Pager window — `1 … 27 28 29`

**Builds on:** Byte 10, Byte 12

**In plain terms:**
29 pages ke liye 29 links dikhana messy hai. Standard trick: hamesha page 1
aur last page, plus current ke aas-paas ke do pages — beech me gap ho to
`…`. Yeh `PagerTokens` function banata hai; `RenderPager` usse clickable
HTML me badalta hai.

**The code:** `default.asp`

```vbs
Function PagerTokens(cur, totalPages)
    Dim bag, ...
    Set bag = Server.CreateObject("Scripting.Dictionary")
    bag.CompareMode = 1
    AddPageToken bag, 1
    AddPageToken bag, totalPages
    ' The two pages either side of the current one - but never past the end...
    For i = cur - 2 To cur + 2
        If i <= totalPages Then AddPageToken bag, i
    Next
    arr = bag.Items
    ' insertion sort - the list is never longer than seven numbers
    ...
    If prev <> 0 Then
        If arr(i) - prev > 1 Then out = out & ",..."
    End If
```

**What's happening:**
`Scripting.Dictionary` (case-insensitive `CompareMode = 1`) duplicate
numbers hata deta hai; 7 se kam items par insertion sort kaafi hai. Har
link `DashHref(p)` se banta hai jo *current* search aur status filter bhi
jodta hai — warna page 2 par click karte hi search gayab ho jati.

**Why it matters:**
"Search stays active when navigating between pages" test isi par khada hai.
`is-active` (current page), `is-disabled` (Previous on page 1) states
screen readers ke liye bhi `aria-current`/`aria-disabled` mark hote hain.

---

## Byte 15: Empty state — filters ke saath alag message

**Builds on:** Byte 12

**In plain terms:**
Jab table khali ho to sirf "No applications yet" dikhana misleading hai agar
user ne koi search kiya ho. Isliye do alag cases: filter active hai to
"No matching applications" + Clear search button, warna asli empty state
jisme briefcase icon aur "Add Application" button hota hai.

**The code:** `default.asp`

```vbs
If rowCount = 0 Then
    If hasFilters Then
        emptyTitle = "No matching applications"
        emptyText  = "Try a different company, role, status or note - or clear the filters to see everything."
    Else
        emptyTitle = "No applications yet"
        emptyText  = "Start by adding your first job or internship application."
    End If
End If
```

**What's happening:**
`hasFilters = (searchRaw <> "" Or statusFilter <> "")`. Markup me
`<div class="empty-state">` ke andar `IconBriefcase()` (pale blue circle
me), title, text aur buttons hain. Nav aur "Add Application" button hamesha
rehte hain — user kahin atka nahi.

**Why it matters:**
Reference design ka sixth screen yahi hai, aur "Show a friendly empty state
when no records exist" spec ki requirement bhi. Dono flavors automated
tests se cover hain (empty table aur no-match search).

---

## Byte 16: Search panel — do search boxes, ek semantics

**Builds on:** Byte 12, Byte 14

**In plain terms:**
Site par do search boxes hain: top bar ka compact wala aur dashboard ka
bada panel (status dropdown ke saath). Dono GET form hain jo
`/default.asp` par submit hote hain — aur *kabhi* `page` field carry
nahi karte, taaki naya search hamesha page 1 se shuru ho.

**The code:** `inc\header.asp`

```asp
<form class="site-search" role="search" method="get" action="/default.asp">
  <input type="search" id="site-search" name="q" maxlength="100"
         placeholder="Search company, role, status, job URL or notes..."
         value="<%= H(searchValue) %>" autocomplete="off">
  <% If filterStatus <> "" Then %>
  <input type="hidden" name="status" value="<%= H(filterStatus) %>">
  <% End If %>
```

**What's happening:**
GET form ka matlab: submit par browser `action?q=...&status=...` bana deta
hai — koi server mutation nahi, isliye bookmark/share bhi safe. `value`
attribute me current query wapas chhapta hai (user dekh sake kya search
kiya tha). `status` hidden field isliye hai taaki top-bar search current
filter todhe na. `app.js` ka clear (×) button field aur form dono ko
dhoondh kar search reset karta hai.

**Why it matters:**
"Submitting a new search starts at page 1" aur "The nav search box keeps
the query" — dono tests isi design par pass hote hain. GET semantics ki
wajah se search links shareable hain.

---

### PUTTING IT TOGETHER (Module 3)

Dashboard ek thin layer hai: query-string inputs ko sanitise karke
`FetchApplicationsPage` deti hai (Module 2), aur result ko render karti
hain — table rows (badges + actions), `results-bar` summary/chips, empty
state ya pager. `header.asp` har page par same top bar aur compact search
deta hai; `styles.css` (reference palette) aur `app.js` (clear button)
presentation ke kaam me CSS/JS me rehte hain. Search-box se row-link tak
ka har dynamic value `H()` se guzra hua hai.

---

# Module 4 — CRUD pages (`new.asp`, `view.asp`, `edit.asp`, `delete.asp`)

## Byte 17: Create — POST/Redirect/GET ka poora chakkar

**Builds on:** Module 1, Module 2

**In plain terms:**
`new.asp` do avasthaon me kaam karta hai: GET par khaali form (CSRF token
ke saath), POST par validation → INSERT → *redirect* details page par.
Redirect isliye ki browser refresh kare to form dobara submit na ho.

**The code:** `new.asp`

```vbs
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    RequireCsrf

    Set values = CollectApplicationValues()
    Set errors = ValidateApplication(values)

    If errors.Count = 0 Then
        isNewId = InsertApplication(values)

        If isNewId > 0 Then
            SetFlash "success", "Application created successfully."
            Response.Redirect "/view.asp?id=" & isNewId
            Response.End
        End If

        ' The insert wrote nothing - treat it as a form error so the values
        ' the user typed are never lost.
        errors("company") = "The application could not be saved. Please try again."
    End If
```

**What's happening:**
`REQUEST_METHOD` se decide hota hai POST ya GET. POST par sabse pehle CSRF
(403 otherwise), phir form values ek dictionary me (`CollectApplicationValues`
— bounded reads), phir validation. Error aaye to page *dobara render* hota
hai: user ki values dobara form me, per-field messages ke saath. Safal
INSERT nayi ID lauta hai → flash + redirect (`Response.Status 302`).

**Why it matters:**
Yeh spec ki "create works with validation" + PRG requirements ka literal
implementation. Insert 0 rows laaye (constraint failure etc.) par user ki
input kabhi lost nahi hoti — wohi form values ke saath wapas aata hai.

---

## Byte 18: Shared form renderer — do pages, ek form

**Builds on:** Byte 17

**In plain terms:**
Create aur edit form identical hain (sirf action URL, button label aur
prefilled values badalte hain). Isliye markup `inc\form.asp` ke
`RenderApplicationForm` me hai — `new.asp` aur `edit.asp` dono ise call
karte hain. Field error helpers (`ErrClass`, `ErrBlock`, `AriaInvalid`)
har field ke neeche message aur `aria-invalid` lagate hain.

**The code:** `inc\form.asp`

```vbs
Sub RenderApplicationForm(action, submitLabel, recordId, v, errs)
...
  <form method="post" action="<%= H(action) %>" class="app-form" novalidate>
    <%= CsrfInput() %>
    <% If recordId <> "" Then %>
    <input type="hidden" name="id" value="<%= H(recordId) %>">
    <% End If %>
    <div class="form-grid">
    <div class="form-field<%= ErrClass(errs, "company") %>">
      <label for="f-company">Company <span class="req" title="Required">*</span></label>
      <input type="text" id="f-company" name="company" maxlength="100"
             placeholder="e.g. Google" autocomplete="organization"
             value="<%= H(v("company")) %>"<%= AriaInvalid(errs, "company") %>>
      <%= ErrBlock(errs, "company") %>
```

**What's happening:**
`novalidate` browser ki built-in validation band karta hai taaki *server*
ke messages dikhen (browser validation bypass karke bhi POST kiya ja sakta
hai). Har field ka `<label for>` input se juda hai — accessibility ka
niyam. Error hone par parent `div` ko `has-error` class milta hai (laal
border) aur `ErrBlock` `<p class="field-error" role="alert">` chhapta hai.

**Why it matters:**
Ek hi form renderer = create aur edit kabhi alag nahi ho sakte (spec:
"share form markup"). A11y (labels, aria-invalid, role=alert) free me aata
hai.

---

## Byte 19: Validation — ek rule set, do pages, aur ek round-trip date check

**Builds on:** Byte 18

**In plain terms:**
`ValidateApplication` saare business rules check karta hai: company/role
2–100 chars, paanch me se ek status, date required (Wishlist chhod kar),
URL http(s) hi ho, notes ≤ 1000. Date check khaas hai: format regex +
`CDate` + wapas ISO me format karke *milana* — `2024-02-30` jaisi rollover
date pakad me aati hai.

**The code:** `inc\validate.asp`

```vbs
Function IsValidIsoDate(s)
    IsValidIsoDate = False
    Dim text
    text = Trim(s & "")
    If Not TestPattern("^[0-9]{4}-[0-9]{2}-[0-9]{2}$", text) Then Exit Function
    ...
    d = CDate(text)
    ...
    If Not IsDate(d) Then Exit Function
    If FormatIsoDate(d) <> text Then Exit Function
    IsValidIsoDate = True
End Function
```

**What's happening:**
Errors `Scripting.Dictionary` me aate hain (field → message); khali dict
matlab valid. Validate karte-karte values *normalise* bhi hoti hain
(`v("applied_on_iso") = appliedIso`) — insert/update ko sirf clean ISO
milta hai. `Wishlist` ka special case: date optional hai, par di ho to
valid honi chahiye. DB ke CHECK constraints (Byte 8 ka last line of
defence) yahan ke rules ko database par bhi enforce karte hain.

**Why it matters:**
"Preserve entered values when validation fails" + "Reject invalid values
with useful messages" dono isi layer se aate hain. Ek rule set hone se
create aur edit par koi alag behaviour nahi.

---

## Byte 20: Details page — id se row tak, ya 404

**Builds on:** Byte 3, Byte 7

**In plain terms:**
`view.asp` teen steps ka hai: id parse karo (`ParseId` → -1 = 404), row
lao (`GetApplication` → `Nothing` = 404), phir saare fields `H()` ke saath
chhapo. Job URL sirf tab link hai jab value ho; warna grey "Not provided".

**The code:** `view.asp`

```vbs
Dim id, rec
id = ParseId(Request.QueryString("id"))
If id < 0 Then NotFound()

Set rec = GetApplication(id)
If rec Is Nothing Then NotFound()
...
<div class="detail-row">
  <dt>Status</dt>
  <dd><%= StatusBadge(rec("status")) %></dd>
</div>
...
  <dd><% If rec("job_url") <> "" Then %><a href="<%= H(rec("job_url")) %>"
      rel="noopener noreferrer" target="_blank"><%= H(rec("job_url")) %></a>
      <% Else %><span class="is-empty">Not provided</span><% End If %></dd>
```

**What's happening:**
`GetApplication` (db.asp) ek dictionary lauta hai `applied_on_iso` aur
`created_at_text` ke saath — formatting data layer me ho chuki hai, page
sirf dikhata hai. `rel="noopener noreferrer"` tab lagta hai jab kisi aur
site ka URL `_blank` me khulta hai (reverse tabnabbing se bachav). Page ke
neeche ek *hidden* POST form hai jiska `form=` attribute Edit/Delete
buttons se juda hai (Byte 22).

**Why it matters:**
"Unknown or non-numeric id → proper 404" yahin implement hai — aur yeh
404 *real status code* ke saath aata hai (Byte 5), koi meta-refresh page
nahi.

---

## Byte 21: Edit — posted id pehle, row count par bharosa nahi

**Builds on:** Byte 17, Byte 20

**In plain terms:**
Edit form `id` ko *hidden field* me bhi bhejta hai. POST par hum pehle
posted id padhte hain, query string ko fallback rakhte hain, aur dono ko
dobara validate karte hain. UPDATE ke baad *rows changed count* check
karta hai — 0 matlab row tab tak delete ho chuki thi → 404.

**The code:** `edit.asp` + `inc\db.asp`

```vbs
    Dim postedId
    postedId = ParseId(Request.Form("id"))
    If postedId < 0 Then postedId = ParseId(Request.QueryString("id"))
    If postedId < 0 Then NotFound()
    ...
        updated = UpdateApplication(id, values)
        If updated > 0 Then
            SetFlash "success", "Application updated successfully."
            Response.Redirect "/view.asp?id=" & id
```

**What's happening:**
`UpdateApplication` row count `cmd.Execute` ke *ByRef* argument se leta hai
(comment me bataya gaya purana bug: `cn.RecordsAffected` property ADO 3001
deti thi). `updated > 0` hi safalta hai; warna ya to row gayab ho chuki thi
(dusre tab me delete) ya database reject ho gaya — dono me user ko honest
jawab milta hai.

**Why it matters:**
"Edit persists every field" aur "unknown id → 404" dono covered. Row count
par bharosaa se "successful-looking" update jo actually kuch badla hi na,
woh possibility khatam.

---

## Byte 22: Delete — POST-only, CSRF, aur modal ka progressive enhancement

**Builds on:** Byte 4, Byte 5, Byte 20

**In plain terms:**
Sabse khatarnak operation sabse zyada suraksha me: GET kabhi delete nahi
karta (405 + `Allow: POST` header), POST par CSRF + id dono check hote
hain. UI me dashboard ki row aur details ki Delete button *asli* POST form
se judi hain; `app.js` sirf ek achha confirmation dialog lagata hai — JS
na ho to browser ka apna `confirm()` chal jata hai.

**The code:** `delete.asp` + `inc\footer.asp` + `static\app.js`

```vbs
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    FailPage 405, "Method not allowed", _
             "Applications can only be deleted from the application list or the details page."
End If
RequireCsrf
Dim id, rec, removed
id = ParseId(Request.Form("id"))
If id < 0 Then NotFound()
...
removed = DeleteApplication(id)
If removed = 0 Then NotFound()
```

```asp
<% ' inc\footer.asp - the shared dialog, rendered on every page %>
<div class="modal-overlay" id="delete-modal" hidden>
  <div class="modal" role="dialog" aria-modal="true"
       aria-labelledby="delete-modal-title" ...>
```

```js
/* static\app.js - intercept clicks on delete triggers */
document.addEventListener('click', function (event) {
  var trigger = event.target.closest ? event.target.closest('.js-delete-trigger') : null;
  if (!trigger) { return; }
  var form = trigger.form || document.getElementById(trigger.getAttribute('form'));
  if (!form || form.method.toLowerCase() !== 'post') { return; }
  event.preventDefault();
  open(form);
}, false);
```

**What's happening:**
Dialog ka `open(form)` target form ko yaad rakhta hai; "Yes, Delete" par
`data-confirmed=true` set karke *bina onsubmit confirm* ke form submit
hota hai (dialog ne jawab le liya). Focus trap aur Esc-close keyboard
users ke liye. Server par wahi purani discipline: method check → CSRF →
`ParseId` → existence check → DELETE → row count (0 = 404) → flash +
redirect dashboard par.

**Why it matters:**
"Never delete data from an ordinary link click" — prefetch, crawler ya
`GET /delete.asp` bookmark kuch delete nahi kar pata (test suite verify
karta hai: row count unchanged). JS sirf polish hai; suraksha hamesha
server par hai.

---

### PUTTING IT TOGETHER (Module 4)

Chaar pages ek hi pattern par chalte hain: **parse inputs → check method
and CSRF → read/validate → call `db.asp` function → flash + redirect (ya
form dobara errors ke saath)**. `new.asp` aur `edit.asp` ka form markup ek
hi `RenderApplicationForm` se aata hai; validation bhi ek hi
`ValidateApplication` hai. `view.asp` aur `delete.asp` "not found" ko
`NotFound()` ke through real 404 dete hain, aur delete par poora defence
stack hai — POST-only method, CSRF token, strict id, row count — jiske
upar `app.js` ka modal sirf ek behtar confirm dialog hai.

---

**Overall picture:** yeh project ek teaching skeleton ki tarah banaya gaya
hai — har concern (encoding, validation, SQL, CSRF, status codes, flash)
ki apni jagah hai, aur har jagah woh code *asli* files me maujood hai jise
smoke suite (177 tests) aur manual checklist dono cover karte hain. Yahi
discipline ASP.NET Core / Flask / FastAPI me middleware, model validation,
ORM aur anti-forgery tokens ke roop me wapas milti hai.
