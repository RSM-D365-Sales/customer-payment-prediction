# Power BI report — build it, then embed it in D365 Finance & Operations

This builds the *embedded analytics* half of the demo. The HTML replica in
`dashboard/` is the pixel-perfect half; this one is the "it's really inside
D365" half. Both read the same workbook, so they always tell the same story.

Budget about 30–40 minutes for the first build. After that, retargeting the
demo at a new prospect is a 2-minute refresh.

---

## Prerequisites

| Need | Why | Who |
|---|---|---|
| Power BI Desktop | Build the report | You |
| Power BI Pro (or PPU / Premium capacity) | Publish + share | You and anyone viewing |
| A Power BI workspace | Host the published report | You |
| D365 F&O with Power BI integration configured | Show the report *inside* F&O | Tenant admin, one-time |

If you can't get the last row done in time, jump to **B4 — no-admin fallback**.
The demo still works.

---

## Part A — Build the report

### A1. Load the workbook

`Home > Get data > Excel workbook` → `data/PaymentPrediction-Demo.xlsx`

Tick these six sheets, then **Transform Data** (not Load):

- `Customers`
- `Transactions`
- `CustomerPredictions`
- `AgedBalances`
- `TopFactors`
- `Config`

Do **not** load `Input_Customers` or `README` — they're the generator's inputs,
not the model.

In Power Query, for each table:

1. `Use First Row as Headers` (usually already applied).
2. Check the data types. The ones that matter:
   - `On time probability`, `Late probability`, `Very late probability` → **Decimal Number**
   - `Balance`, `Amount in transaction currency`, all `Aged*` columns → **Decimal Number**
   - `Date`, `Due date`, `Last payment date`, `As of date` → **Date**
   - `Days past due`, all count columns → **Whole Number**
   - `Config[Value]` → leave as **Text** (it holds mixed types by design)
3. `Close & Apply`.

### A2. Relationships

Model view. You want a single-direction star from `Customers`:

| From | To | Cardinality |
|---|---|---|
| `Customers[Customer account]` | `Transactions[Customer account]` | 1 → * |
| `Customers[Customer account]` | `CustomerPredictions[Customer account]` | 1 → 1 |
| `Customers[Customer account]` | `AgedBalances[Customer account]` | 1 → 1 |
| `Customers[Customer account]` | `TopFactors[Customer account]` | 1 → * |

`Config` stays disconnected — the measures reach into it with `LOOKUPVALUE`.

Hide the duplicate `Name` column on every table except `Customers`, so the
field list stays clean.

### A3. Apply the theme

`View > Themes > Browse for themes` → `powerbi/d365-payment-prediction-theme.json`

This is what makes the report stop looking like Power BI: Segoe UI throughout,
square corners, 1px `#E1DFDD` borders, no visual headers, and the D365 green /
orange / red for on time / late / very late.

### A4. Add the measures

Create an empty table to hold them: `Home > Enter data`, name it `_Measures`,
one column, no rows, Load. Then paste each block from `powerbi/measures.dax` via
`Home > New measure`.

Format as you go:
- `On time %`, `Late %`, `Very late %`, `Predicted late %`, `Credit used %`,
  `Aged 90 plus %` → **Percentage, 0 decimals**
- every `* amount`, `Open amount`, `Credit limit` → **Decimal, 2 decimals, thousands separator**

### A5. Page 1 — "Payment predictions per customer"

Canvas: `View > Page view > Actual size`, page size **1280 × 720**.
This is the aspect ratio that sits cleanly in a D365 Analytics tab.

| Visual | Position (x, y, w, h) | Fields |
|---|---|---|
| Text box — `Payment predictions per customer`, 16pt Segoe UI Semibold | 20, 14, 460, 30 | — |
| Slicer (dropdown) | 950, 14, 310, 34 | `Customers[Customer group]` |
| Card | 20, 58, 226, 104 | `[On time amount]` |
| Card | 258, 58, 226, 104 | `[Late amount]` |
| Card | 496, 58, 226, 104 | `[Very late amount]` |
| Card | 734, 58, 226, 104 | `[Open amount]` |
| Card | 972, 58, 288, 104 | `[At risk transactions]` |
| Table | 20, 176, 700, 528 | see below |
| 100% stacked bar | 732, 176, 528, 260 | Axis `Customers[Customer group]`, Values `[On time %]`, `[Late %]`, `[Very late %]` |
| Clustered column | 732, 448, 528, 256 | Axis: the five `Aged *` measures |

Table columns, in this order — it mirrors the real F&O list page:

`Customer account`, `Name`, `Balance`, `Currency`, `[On time %]`, `[Late %]`, `[Very late %]`

Then the detail that sells it — **conditional formatting on `[On time %]`**:

> Right-click the `On time %` field in the Values well →
> `Conditional formatting > Font colour` →
> Format style: **Field value** → Based on field: **`[On time colour]`** → OK.

Now anything under 40% renders red, 40–60% orange, above that green. Same
thresholds the D365 grid uses for its red dot.

### A6. Page 2 — "Transactions predicted to be paid late"

Duplicate page 1, retitle it, and swap the table for the transaction grid:

`Customer account`, `Name`, `[Risk dot]`, `[On time %]`, `Voucher`, `Invoice`,
`Transaction type`, `Date`, `Due date`, `Amount in transaction currency`,
`Balance`, `Currency`

Set the `[Risk dot]` column width to ~24px — it renders a `●` only below the
threshold, exactly like the F&O grid. Apply the same `[On time colour]`
conditional format to it and to `[On time %]`.

Replace the two charts on the right with three **Gauge** visuals stacked
vertically (`[On time %]`, `[Late %]`, `[Very late %]`, each Max value = 1) and a
**Multi-row card** underneath bound to `[Open invoices]`, `[Late invoices]`,
`[Open collection cases]`, `[Open activities]`, `[Average days to pay]`.

Set the gauge fill colours manually — green `#107C10`, orange `#E36C0A`,
red `#C50F1F` — because the theme can only carry one default gauge colour.

Because everything cross-filters, clicking a row drives the gauges and counters
the way the D365 Related-information pane does.

---

## Part B — Embed in D365 F&O

### B1. Publish

`Home > Publish` → pick the workspace. Confirm it opens in the Power BI Service
and renders correctly there first. If it looks wrong in the Service, it will
look wrong in D365.

### B2. Configure Power BI integration (one-time, needs an admin)

This is the gate. In F&O:

`System administration > Setup > Power BI > Power BI configuration`

It needs an Entra ID (Azure AD) app registration with delegated Power BI
Service permissions, and the Application ID + secret entered here. Your tenant
admin does this once; after that every F&O user can surface Power BI content.

> Menu paths shift between 10.0.x releases. If `Power BI configuration` isn't
> where this says, search the F&O navigation for "Power BI" — it has moved
> between `System administration` and `System parameters` across versions.
> Confirm the exact path in your demo environment before the meeting.

### B3. Put the report on a workspace

1. Open the workspace you want to demo from — **Credit and collections
   management** is the natural home for payment predictions.
2. `Options > Personalize this form` (or click the **Analytics** tab if the
   workspace already has one).
3. `Insert > Power BI report`.
4. Choose your Power BI workspace, then the report, then the page.
5. Save the personalization. To make it stick for other demo users, use
   `Options > Personalize > Manage > Publish` and target a user or role.

Result: a tab inside F&O, wrapped in real D365 chrome, showing your report.

### B4. No-admin fallback

If B2 can't happen before your demo, either of these still lands:

- **Power BI Service full-screen.** Open the report in the Service, press
  focus/full-screen. Alt-tab from F&O to it. Slightly less seamless, zero setup.
- **Use the HTML replica instead.** `dashboard/payment-prediction.html` needs no
  tenant, no licence, no network. It's the higher-fidelity option anyway — it
  reproduces the Related-information fly-out and the three-arc gauges that
  Power BI visuals can only approximate.

Honest recommendation: **lead with the HTML replica for the UI story, and bring
the Power BI report out when the conversation turns to "and here's how you'd
extend it with your own analytics."** They answer different objections.

---

## Part C — Retarget for the next demo

1. Edit `Input_Customers` in `data/PaymentPrediction-Demo.xlsx` — replace the
   rows with the prospect's customers or tenants.
2. Update `Config` → `Demo name`, `Legal entity`, `Organization`, `As of date`.
3. Run `python scripts/build_demo.py`.
4. In Power BI Desktop: `Home > Refresh`. Republish.

The schema never changes, so no visual, measure, or relationship breaks.

---

## Part D — Moving to Fabric later

The workbook schema was designed to lift into Fabric unchanged. When you want
it there:

1. Land the five output sheets as Delta tables in a Lakehouse (or tables in a
   Warehouse) — same table names, same column names.
2. In Power BI, swap the Excel source for the Lakehouse via
   `Transform data > Data source settings`, or rebuild the model as Direct Lake.
3. Every measure in `measures.dax` keeps working — they only reference table and
   column names, never the connector.

Worth doing if you want the demo dataset shared across a team or refreshed on a
schedule. Not worth doing for a single-presenter demo — the Excel file is
faster to change in the ten minutes before a meeting.
