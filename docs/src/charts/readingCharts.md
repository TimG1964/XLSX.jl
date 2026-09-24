```@meta
CurrentModule = XLSX.Charts
```

!!! note "Experimental"

    Handling of native Excel charts in XLSX.jl is experimental and a work in 
    progress. All aspects may be subject to change. Feedback is welcome!


# Reading charts

```julia
julia> using XLSX, XLSX.Charts
```

This page is about getting data and metadata out of a chart: its title, its
series, where each series plots from, and the values Excel cached inside the
chart part.

## The chart cache

Every `c:` chart part carries a snapshot of the data it plots, written by Excel
each time the file is saved. That is what makes a chart render when its source
is unavailable — the sheet has been deleted, or the data lives in a workbook
that isn't to hand.

The cache is what XLSX.jl reads. It is never recomputed from the worksheet, so
two things follow:

- The values reflect the source cells **at the last save by Excel**. If the
  source has changed since, the cache is stale.
- A file written by a tool that doesn't populate the cache has charts with no
  cached values at all, even though the chart is valid and displays correctly
  once Excel opens and re-saves it.

Where a series refers to a real range, [`getChartRanges`](@ref)
gives you that range so you can read the live cells instead. See
[Where the data came from](@ref).

`cx:` charts carry no cache at all — Excel writes references only — so this
whole section applies to [`Chart`](@ref) and not to [`ChartEx`](@ref).

## Title and series

```julia
julia> f = XLSX.readxlsx("chart_basic.xlsx");

julia> c = getChart(f["Data"], "chart1");

julia> getChartTitle(c)
"Revenue by Region"

julia> ss = getChartSeries(c)
2-element Vector{XLSX.Charts.ChartSeries}:
 XLSX.Charts.ChartSeries 2024 (barChart)
 XLSX.Charts.ChartSeries 2025 (barChart)
```

A [`ChartSeries`](@ref XLSX.Charts.ChartSeries) holds its `idx` and `order` as
Excel recorded them, its `charttype`, its `name`, and up to four references:
`name_ref`, `categories`, `values` and `bubble_sizes`.

```julia
julia> s = ss[1];

julia> s.values
XLSX.Charts.ChartRef Data!$B$2:$B$5 (num, 4 pts)
  format: General
  data: [10.0, 20.0, 15.0, 5.0]

julia> s.values.ref
"Data!\$B\$2:\$B\$5"

julia> s.values.kind
:num
```

!!! note
    `categories` holds `c:cat` for category charts and `c:xVal` for scatter and
    bubble charts; `values` holds `c:val` or `c:yVal` correspondingly. The two
    fields mean the same thing whatever the chart type, so you never need to
    branch on `charttype` to find the data.

A [`ChartRef`](@ref)'s `kind` is one of `:num` (numbers from cells, `c:numRef`),
`:str` (text from cells, `c:strRef`), `:multiLvlStr` (text from cells on several
category levels, `c:multiLvlStrRef`), `:numLit` (numbers typed into the chart,
`c:numLit`) or `:strLit` (text typed into the chart, `c:strLit`). The two `Lit`
kinds have no `ref`, since they read from no cells.

A chart's groups — what Excel calls its chart types — come from
[`getChartTypes`](@ref), one entry per group in the
plot area:

```julia
julia> getChartTypes(c)
1-element Vector{Symbol}:
 :barChart
```

## Getting the cached data

[`getChartData`](@ref) flattens a chart's series into
an [`XLSX.DataTable`](@ref), ready for `DataFrame` or any other Tables.jl sink:

```julia
julia> using DataFrames

julia> DataFrame(getChartData(c))
4×3 DataFrame
 Row │ categories  2024     2025
     │ String      Float64  Float64
─────┼──────────────────────────────
   1 │ North          10.0     12.0
   2 │ South          20.0     18.0
   3 │ East           15.0     25.0
   4 │ West            5.0      9.0
```

It also takes a chart name directly, saving the intermediate `getChart` call:

```julia
julia> dt = getChartData(f["Data"], "chart1");
```

Series with no name in the file — the ones Excel labels "Series1", "Series2" in
the legend — are labelled by position. Duplicate names get a numeric suffix.
Series of unequal length are padded with `missing`.

### How the columns are laid out

Where every series shares one category reference, a single `categories` column
leads the table. On a scatter or bubble chart the categories are the x values,
so that column is numeric rather than text, and a bubble chart adds a
`<series>_size` column after each series' values:

```julia
julia> DataFrame(getChartData(XLSX.readxlsx("chart_kinds.xlsx"), "chart2"))
4×3 DataFrame
 Row │ categories  Series1  Series1_size 
     │ Float64     Float64  Float64      
─────┼───────────────────────────────────
   1 │        1.0     10.0           3.0
   2 │        2.0     20.0           5.0
   3 │        3.0     15.0           2.0
   4 │        4.0      5.0           4.0
```

Where the series don't share a category reference — which a scatter or bubble
chart allows, each series carrying its own x values — each contributes its own
`<series>_x` column immediately before its values instead.

```julia
julia> DataFrame(getChartData(XLSX.readxlsx("chart_scatter_xy.xlsx"), "chart1"))
4×4 DataFrame
 Row │ y1_x     y1       y2_x     y2      
     │ Float64  Float64  Float64  Float64 
─────┼────────────────────────────────────
   1 │     1.0     10.0      2.0      5.0
   2 │     2.0     20.0      4.0     15.0
   3 │     3.0     30.0      6.0     25.0
   4 │     4.0     40.0      8.0     35.0
```

Multi-level categories give one column per level, in the order Excel writes them
(innermost first).

No category column is produced when the chart has no `c:cat` at all: Excel
plots against an implicit index in that case, and caches nothing for it.

## Where the data came from

The cache tells you what the chart *displayed*;
[`getChartRanges`](@ref) tells you where it says the
data *came from*, as a range you can hand straight back to
[`XLSX.getdata`](@ref):

```julia
julia> r = getChartRanges(f["Data"], "chart1");

julia> r[1].name, r[1].categories, r[1].values
("2024", Data!A2:A5, Data!B2:B5)

julia> XLSX.getdata(f, r[1].values)      # the live cells, not the cache
4×1 Matrix{Any}:
 10
 20
 15
  5
```

One entry per series, in document order, parallel to `getChartSeries(c)`. Each
carries the series `idx` and `name` alongside its `categories`, `values` and
`bubble_sizes` ranges. `bubble_sizes` is `nothing` for every chart type but
bubble.

Called with no name, `getChartRanges` covers every chart on the sheet or in the
workbook, each paired with its chart name.

A range is `nothing` wherever the series has no addressable source: a literal
series, or a reference to an external workbook.

!!! note
    A chart may plot from a range whose sheet no longer exists, or which has
    since been overwritten with something else. The range records the chart's
    claim about its source, not a guarantee about the current contents of those
    cells. Where the two disagree, the cache is the older of the pair.

## Blanks and errors

Excel writes the cache as a sparse list of points against a declared `ptCount`.
A blank cell in the source is omitted from that list, so it arrives as `missing`
in `data`.

Errors are narrower. `#N/A` is cached as the literal string `#N/A`, and XLSX.jl
records it in `errors` and stores `missing` in its place — so it is not mistaken
for a number, and can be told apart from a blank. **Every other error value is
written into the cache as `0`.** By the time the file is on disk, a `#DIV/0!`
and a genuine zero are the same three bytes.

[`XLSX.iserror`](@ref) and [`XLSX.geterror`](@ref) report what the cache
preserved, which is `#N/A` and nothing else:

```julia
julia> g = getChart(XLSX.readxlsx("chart_gaps.xlsx"), "chart1");

julia> v = getChartSeries(g)[1].values
XLSX.Charts.ChartRef Data!$B$2:$B$6 (num, 5 pts)
  format: General
  errors at: 3
  data: Union{Missing, Float64}[1.0, missing, missing, 0.0, 5.0]

julia> XLSX.iserror(v)
5-element Vector{Bool}:
 0
 0
 1
 0
 0

julia> XLSX.geterror(v)
5-element Vector{String}:
 ""
 ""
 "#N/A"
 ""
 ""
```

The source column behind this chart shows all three outcomes at once:

| Cell | Source | In the cache | In `data` |
| --- | --- | --- | --- |
| `B2` | `1` | `<c:pt idx="0">1</c:pt>` | `1.0` |
| `B3` | *(blank)* | omitted | `missing` |
| `B4` | `=NA()` → `#N/A` | `<c:pt idx="2">#N/A</c:pt>` | `missing`, flagged in `errors` |
| `B5` | `=1/0` → `#DIV/0!` | `<c:pt idx="3">0</c:pt>` | `0.0` |
| `B6` | `5` | `<c:pt idx="4">5</c:pt>` | `5.0` |

!!! warning
    A zero in cached chart data may be a genuine zero or may be any error other
    than `#N/A`. This is a limitation of the file format, not of XLSX.jl: Excel
    discards the distinction when it writes the cache. Where it matters, read
    the source cells through [`getChartRanges`](@ref)
    and [`XLSX.getdata`](@ref), which see the real cell values and report every
    error type.

!!! note
    [`XLSX.iserror`](@ref) and [`XLSX.geterror`](@ref) on a cell range report
    every error value. On a `ChartRef` they can only report what survived into
    the cache. The functions behave the same way; the data they are given does
    not.

## Combo charts

A chart whose plot area holds more than one group — a bar series and a line
series sharing an axis, say — reports every group, and each series remembers
which group it belongs to:

```julia
julia> cc = getChart(XLSX.readxlsx("chart_combo.xlsx"), "chart1");

julia> getChartTypes(cc)
2-element Vector{Symbol}:
 :barChart
 :lineChart

julia> [(s.name, s.charttype) for s in getChartSeries(cc)]
2-element Vector{Tuple{Union{Nothing, String}, Symbol}}:
 ("2024", :barChart)
 ("2025", :lineChart)
```

[`getChartData`](@ref) is indifferent to this: series from every group land in the same
table, in document order.

## Charts referring to another workbook

A series may plot from a range in a different workbook, in which case its `ref`
begins with a bracketed index into the workbook's external references:

```julia
julia> e = getChart(XLSX.readxlsx("chart_external.xlsx"), "chart1");

julia> getChartSeries(e)[1].values.ref
"[1]Feuil1!\$A\$1:\$A\$10"
```

Pass `get_external_refs=true` to substitute the recorded workbook path, as
[`XLSX.getFormula`](@ref) does for formulas:

```julia
julia> getChartSeries(e; get_external_refs=true)[1].values.ref
"[Test2.xlsx]Feuil1!\$A\$1:\$A\$10"
```

The cached values are available either way — that is the point of the cache —
but [`getChartRanges`](@ref) returns `nothing` for such a series, since the range is not
addressable within this workbook.

## Reading metadata only

Parsing the cached points is the expensive part of reading a large chart. Where
you only need the shape of a chart — its title, types, series names, source
formulas, format codes and point counts — pass `read_cached_values=false`:

```julia
julia> s = getChartSeries(c; read_cached_values=false)[1];

julia> s.values
XLSX.Charts.ChartRef Data!$B$2:$B$5 (num, 4 pts)
  format: General

julia> s.values.data
Any[]
```

`ptCount` is still populated; only `data` is left empty. Series names are always
read, since they are metadata rather than plotted values.

[`getChartRanges`](@ref) uses `read_cached_values=false` internally, as it never needs
the values.

## Reading a chartEx chart

A [`ChartEx`](@ref) reads its title, type and source ranges the same way:

```julia
julia> x = getCharts(XLSX.readxlsx("chartex_layouts.xlsx"))[1];

julia> getChartType(x), getChartTitle(x)
(:waterfall, "Cash flow")

julia> getChartRanges(x)
1-element Vector{Union{Nothing, XLSX.NonContiguousRange, XLSX.SheetCellRange, XLSX.SheetCellRef, XLSX.SheetColumnRange, XLSX.SheetRowRange}}:
 waterfall!A1:A5

julia> x = getCharts(XLSX.readxlsx("chartex_formatted.xlsx"))[1];

julia> getChartType(x), getChartTitle(x)
(:waterfall, "Custom title")
```

There are no cached values to read, so [`getChartData`](@ref) throws for a [`ChartEx`](@ref).
Its references are written indirectly, through hidden defined names of the form
`_xlchart.v1.0`, which [`getChartRanges`](@ref) resolves; those names are excluded from
[`XLSX.getDefinedNames`](@ref) and are protected from deletion.

`cx:` charts have their own data model — dimensions grouped into data blocks —
reached through [`getChartDataBlocks`](@ref) and
[`getSeriesData`](@ref).
