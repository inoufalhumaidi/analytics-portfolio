# Power BI Desktop Build Guide — Ridgeline Field Services

Produces `Ridgeline_KPI_Dashboard.pbix`. Power BI Desktop has no CLI, but it does expose its data model (not its report pages) through a local Analysis Services engine while a file is open — the same mechanism Tabular Editor and DAX Studio use. Two ways to build this:

**Option A — scripted (recommended, ~2 minutes for the data model):** run `powerbi/Build_PowerBI_Model.ps1` against an open blank report. It connects to Power BI Desktop's local engine and builds all 7 tables, 6 relationships (5 active), all 31 DAX measures and the Year › MonthName hierarchy by script — Sections 1–4 below, done for you. You still build the report pages by hand in Section 5. The script reaches only the data model, not report layout.

```powershell
# 1. Open Power BI Desktop, File > New (a blank report)
# 2. Run sql/01_create_schema.sql through sql/02_generate_synthetic_data.sql if you haven't
# 3. Run this from the powerbi/ folder:
powershell -File Build_PowerBI_Model.ps1
```

Power BI Desktop will show a credentials/privacy-level dialog the first time it processes the new SQL Server connection — the script will appear to hang until you click through it (choose Windows/Integrated auth, any privacy level). This happens once per Desktop session. Once it finishes, expand `Fact_ServiceJobs` in the Fields pane — the 31 measures are grouped into Volume/Quality/Utilization/Financial/Targets/Status/Formatting folders underneath it, not in a separate table. Save the file (File > Save As) into this folder when done, then skip to Section 5.

**Option B — fully manual (~15–20 minutes):** follow Sections 1–4 below yourself in the Desktop UI, then Section 5 either way.

**Prerequisite (both options):** run `sql/01_create_schema.sql` through `sql/05_stored_procedures.sql` against your SQL Server instance first (see `/README.md`).

## 1. Connect to the data

1. Open Power BI Desktop → **Get Data** → **SQL Server**.
2. Server: `localhost\TEW_SQLEXPRESS` (or your instance name). Database: `RidgelineFieldOps`.
3. Data Connectivity mode: **Import** (this dataset is small enough that DirectQuery isn't necessary, and Import gives faster visuals).
4. In the Navigator, check these objects and click **Load**:
   - `Dim_Date`, `Dim_Region`, `Dim_Technician`, `Dim_Client`, `Dim_ServiceType`, `Ref_KPITargets`
   - `Fact_ServiceJobs`
   - (Do not load the `vw_*` views or `usp_*` procs — the whole point of this model is to rebuild the KPI logic natively in DAX so it's inspectable inside the .pbix; the SQL views remain the reusable/auditable reference implementation.)

## 2. Fix data types on load (Power Query)

In Power Query Editor, before loading:
- `Dim_Date[Date]` → Date
- `Fact_ServiceJobs[ScheduledStart]`, `[ActualStart]`, `[ActualEnd]` → Date/Time
- `Fact_ServiceJobs[FirstTimeFix]`, `[CallbackFlag]`, `[SLAMet]` → Whole Number (they arrive as SQL `bit`; Power BI sometimes maps this to True/False — convert to 0/1 so `AVERAGE()` works as documented in `DAX_Measures.md`)

## 3. Build relationships (Model view)

Power BI will likely auto-detect these from the FK constraints, but verify each is **Single** direction, **Many-to-one** (Fact → Dim):

| From | To | Active? |
|---|---|---|
| `Fact_ServiceJobs[DateKey]` | `Dim_Date[DateKey]` | Yes |
| `Fact_ServiceJobs[TechnicianKey]` | `Dim_Technician[TechnicianKey]` | Yes |
| `Fact_ServiceJobs[ClientKey]` | `Dim_Client[ClientKey]` | Yes |
| `Fact_ServiceJobs[ServiceTypeKey]` | `Dim_ServiceType[ServiceTypeKey]` | Yes |
| `Dim_Technician[RegionKey]` | `Dim_Region[RegionKey]` | **Yes** |
| `Fact_ServiceJobs[RegionKey]` | `Dim_Region[RegionKey]` | **No — deactivate** |

**Region is deliberately a snowflake (`Dim_Region → Dim_Technician → Fact`), not a direct star spoke.** Power BI auto-detects the direct `Fact_ServiceJobs[RegionKey] → Dim_Region[RegionKey]` relationship, and that one is a trap for this model: capacity measures iterate `Dim_Technician`, and a filter on `Dim_Region` reaches only the fact table — filters do not travel dimension → fact → dimension. With the direct path active, slicing a visual by `Dim_Region[RegionName]` narrows `[Worked Minutes]` but leaves `[Available Minutes]` at the full-company total, so regional `Utilization %` comes out silently understated (roughly a fifth of its true value) with no error to warn you. Routing region through the technician fixes it, and leaves no ambiguous second path. This is safe here only because every job's `RegionKey` equals its technician's `RegionKey` — confirmed with `SELECT COUNT(*) FROM Fact_ServiceJobs f JOIN Dim_Technician t ON t.TechnicianKey = f.TechnicianKey WHERE f.RegionKey <> t.RegionKey` returning 0.

Mark `Dim_Date` as a **Date table** (right-click table → Mark as date table → pick `Date` column) so time-intelligence functions work if you add any later.

## 4. Add the measures table

*(Skip this section if you used `Build_PowerBI_Model.ps1` — it adds all measures directly onto `Fact_ServiceJobs`, organized into folders, instead of a separate `_Measures` table.)*

Modeling → New Table:
```
_Measures = ROW("placeholder", 0)
```
Delete the `placeholder` column afterward (Power BI won't let a table be truly empty, this is the standard workaround). Paste every measure from `DAX_Measures.md` into `_Measures` (New Measure, one at a time, or paste the DAX into the formula bar for each).

## 5. Build report pages

**Page 1 — Regional Overview**
- Card visuals across the top: `[Total Jobs]`, `[Utilization %]`, `[First-Time-Fix %]`, `[SLA Compliance %]`
- Clustered column chart: `Dim_Region[RegionName]` (axis) × `[Utilization %]` (value) — this is the chart that shows South Metro's underutilization at a glance
- Line chart: shows the HVAC/Electrical seasonality
  - **X-axis:** `Dim_Date` › `Year/Month Hierarchy`, *not* `Dim_Date[Date]`. This model has no automatic date hierarchy, so `Date` has no Month level and plots ~730 daily points.
  - **Y-axis:** `[Total Jobs]`
  - **Legend:** `Dim_ServiceType[ServiceCategory]`
  - Then use the visual's **"Expand all down one level in the hierarchy"** button (the forked double-arrow above the chart) so the axis reads *2024 Jan … 2025 Dec*, 24 points. Plain "Drill down" instead shows only one year's months at a time, and the Month level on its own merges both years into 12 points and hides the trend.
  - Correct result: two lines, 24 points each, with HVAC peaking in Jun–Aug and Dec–Feb and Electrical higher in the months between.
- **Two separate slicers**, never one:
  - Slicer 1: `Dim_Region[RegionName]` only
  - Slicer 2: `Dim_Date[Year/Month Hierarchy]` (Year → Month)

  Putting `RegionName` and the Year/Month levels in the *same* slicer fails with `DataViewMappingError_ConditionNotHierarchicallyRelated`. A slicer hierarchy needs a one-to-many chain from each level to the next. `Dim_Region` and `Dim_Date` only meet through `Fact_ServiceJobs`, and they reach it from opposite sides, so no chain exists between them. Levels taken from a single table, such as Year → Month, always work.

  For month *names* in any visual, use `Dim_Date[MonthName]`; it is sorted by `Month` so it runs January to December instead of alphabetically. Leave `Dim_Date[Date]` with no sort-by column. Sorting it by `Month` puts every January from both years ahead of February and breaks the line chart below.

**Page 2 — Technician Priority Action Queue**
- Table visual: `Dim_Technician[TechnicianName]`, `Dim_Region[RegionName]`, `Dim_Technician[SkillLevel]`, `[Utilization %]`, `[First-Time-Fix %]`, `[SLA Compliance %]`, `[Callback %]`, `[Risk Score]`
- Sort by `[Risk Score]` descending; add a visual-level filter `[Risk Score] > 0`
- **Filter this page by the Page 1 slicers.** Without a date filter the table scores each technician over the full two years (20 flagged) instead of the month being reviewed (25 flagged in December 2025). Sync both Page 1 slicers here: *View → Sync slicers*, select the slicer on Regional Overview, tick **Sync** for this page and leave **Visible** unticked. Do the same for Page 3. The shipped `.pbix` does this with hidden copies of both slicers (sync groups `RegionName` and `Year/Month Hierarchy`), so pages 2 and 3 always follow the selection made on Regional Overview. Pages 2 and 3 don't show that selection, so set it on Page 1 before reading them.
- Conditional formatting — one rule per KPI column, all with **Background color** → *fx* → *Format style: Field value*:

  | Column (series) | Based on field |
  |---|---|
  | `Utilization %` | `Utilization Color` |
  | `First-Time-Fix %` | `FTF Color` |
  | `SLA Compliance %` | `SLA Color` |
  | `Callback %` | `Callback Color` |

  Path: select the table → *Format* → *Cell elements* → choose the series → turn on *Background color* → click *fx*. The colour measures are in the **Formatting** folder under `Fact_ServiceJobs`.
  - Don't use the `... Status` measures here: they return text, and *Field value* formatting needs a colour, so a text measure applies no colour at all.
  - Don't set Font color to the same measure as well, or the text will match its cell and disappear.
  - Cells turn red only on a threshold breach (orange for utilization over capacity); everything else stays unformatted. Each coloured cell corresponds to one point of that row's `Risk Score`.

**Page 3 — SLA & Quality Drilldown**
- Matrix: rows = `Dim_Region[RegionName]`, columns = `Dim_ServiceType[ServiceName]`, values = `[SLA Compliance %]` — this is where the East Valley urgent-job SLA gap becomes visible (filter to `Dim_ServiceType[SLAHours] = 1` to isolate the tight-window services)
- Scatter chart, one dot per technician — *are stretched technicians making more mistakes?*
  - **Values (Details):** `Dim_Technician[TechnicianName]`
  - **Legend:** `Dim_Technician[SkillLevel]`
  - **X-axis:** `[Utilization %]`
  - **Y-axis:** `[First-Time-Fix %]`
  - **Size:** `[Total Jobs]`

  Don't plot `[First-Time-Fix %]` against `[Callback %]`. A job without a first-time fix *is* a callback, so the two always add to exactly 100% (confirmed on all 33,700 completed jobs), and the points can only land on one straight line. An earlier version of this guide made that mistake.

### Navigation bar (all pages)

Each page has three buttons: Regional Overview, Technician Priority Action Queue, SLA & Quality Drilldown. For **every** button, including the one pointing at the page it sits on:

1. Select the button → *Format button* → **Action** → switch **On**.
2. **Type:** Page navigation.
3. **Destination:** the page named on the button.

The usual reason a button "does nothing" is that the destination is set but **Action is still Off**. In Power BI Desktop, **Ctrl + click** a button to follow it while editing. A plain click only works in Reading view or after publishing.

## 6. Publish (optional — requires your own Power BI account/license)

File → Publish → Publish to Power BI, choose a workspace. This is what makes the report genuinely "live" on the web instead of a local file — it's optional and outside this repo's scope since it depends on your account, but the report is built to support it with no further changes.

## 7. Save

File → Save As → `Ridgeline_KPI_Dashboard.pbix`, save into this `powerbi/` folder before committing to the repo.
