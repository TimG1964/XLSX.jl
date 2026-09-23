# Charts

(Chart support in XLSX.jl Work in Progress)

Excel charts come in two schemas, and XLSX.jl reads, formats and creates both.

- The **`c:` schema** (ECMA-376) covers bar, column, line, pie, doughnut, area,
  scatter, bubble, radar and stock charts — everything Excel has had for a long
  time. These are [`XLSX.Charts.Chart`](@ref).
- The **`cx:` schema** ("chartEx", a Microsoft extension) covers waterfall,
  funnel, treemap, sunburst, histogram, Pareto, box & whisker and region map.
  These are [`XLSX.Charts.ChartEx`](@ref).

Chart support lives in its own sub-module. Qualify the functions as
`XLSX.Charts.getCharts`, or bring them into scope:

```julia
julia> using XLSX, XLSX.Charts
```

The rest of this guide uses the unqualified form. The handle and value types are
not exported, so they stay qualified: `XLSX.Charts.ChartSeries`.

This page covers the ideas the rest of the guide assumes. The other pages are
[Reading charts](readingCharts.md), [Formatting charts](formattingCharts.md),
[Creating charts](creatingCharts.md), [Creating chartEx charts](creatingChartEx.md),
[Sheets with charts](chartSheets.md) and [Limitations](chartLimitations.md).

## Finding the charts

[`getCharts`](@ref) lists the charts anchored to a worksheet, or every chart in
the workbook:

```julia
julia> f = XLSX.readxlsx("chart_basic.xlsx");

julia> getCharts(f["Data"])          # this sheet
1-element Vector{XLSX.Charts.Chart}:
 XLSX.Charts.Chart("chart1", "Data", G9, barChart, 2 series)

julia> getCharts(f)                  # the whole workbook
1-element Vector{XLSX.Charts.Chart}:
 XLSX.Charts.Chart("chart1", "Data", G9, barChart, 2 series)
```

`getCharts(ws)` returns them in the order the drawing declares them.
`getCharts(xf)` returns those for every sheet in sheet order, followed by any
chart part the package declares that no drawing references — an orphan left
behind by an edit elsewhere. Both give an empty vector when there is nothing to
find.

[`getChart`](@ref) fetches one, by part name with or without its extension, by
package path, or by relationship id within the drawing:

```julia
julia> c = getChart(f["Data"], "chart1")
XLSX.Charts.Chart "chart1" on sheet "Data" at G9:N23
  title: "Revenue by Region"
  type: barChart
  series: 2
    [1] 2024 - Data!$B$2:$B$5 (4 pts)
    [2] 2025 - Data!$C$2:$C$5 (4 pts)

julia> getChart(f, "xl/charts/chart1.xml") == c
true

julia> getChart(f["Data"], "rId1") == c
true
```

Asking for a chart that isn't there throws an `XLSXError` listing the ones that
are.

## Which schema is this?

A workbook can hold both kinds, so `getCharts` returns a vector that may mix
`Chart` and `ChartEx`. [`getChartSchema`](@ref) tells them apart, and
[`getChartType`](@ref) names the kind:

```julia
julia> getChartSchema(c), getChartType(c)
(:c, :barChart)
```

Most of the API takes either. Where it doesn't, it's because the schemas
genuinely differ — a `cx:` chart has no cached values, for instance — and the
guide says so.

## Handles never go stale

A `Chart` is a handle: it holds the package and the part's path, and nothing
else. Every accessor reads the part as it stands at that moment, and every
setter returns the handle unchanged.

```julia
julia> setChartTitleText(c, "Revenue 2024-25") === c
true
```

So a handle read before a write still works afterwards, and setters chain
naturally. You never need to re-fetch a chart after modifying it.

## Values carry keys

Everything read *from* a chart — a series, an axis, a group, a data point, a
label, a trendline, error bars, a marker — is a value rather than a handle. Each
carries the key that identifies its element:

| Value | Key |
| --- | --- |
| [`ChartSeries`](@ref XLSX.Charts.ChartSeries) | `c:idx` |
| [`ChartAxis`](@ref XLSX.Charts.ChartAxis) | `c:axId` |
| [`ChartGroup`](@ref XLSX.Charts.ChartGroup) | its tag and axis ids |
| [`ChartDataPoint`](@ref XLSX.Charts.ChartDataPoint), [`ChartDataLabel`](@ref XLSX.Charts.ChartDataLabel) | owning series and point index |
| [`ChartTrendline`](@ref XLSX.Charts.ChartTrendline), [`ChartErrorBars`](@ref XLSX.Charts.ChartErrorBars) | owning series and ordinal |

Any function taking `(c, x)` finds `x`'s element by that key in the current
part. So a value read before a write still addresses the right element
afterwards, even though its other fields describe the part as it was when read:

```julia
julia> ax = getChartAxes(c, :value)[1];

julia> setSeriesFill(c, 1, "FF0000");        # unrelated write

julia> getAxisMajorUnit(c, ax)               # `ax` still resolves
5.0
```

Re-read the value when you want its own fields refreshed; keep using it as an
argument either way.

## Positions count from 1

A series position, and a data point position, count from 1 as everywhere else in
the package. Excel's own identifiers are something else:

```julia
julia> [s.idx for s in getChartSeries(c)]    # c:idx, as the file records it
2-element Vector{Int64}:
 0
 1
```

They usually agree, because Excel numbers series from 0 in order. They need not:
a file where series were deleted can have gaps. Positions are resolved in
document order on each call, so they are always what you see in
`getChartSeries(c)`.

## Absent is not the same as off

Optional values are `Union{Nothing,T}`, and `nothing` means the file says
nothing — so Excel inherits from the chart style. It never means "off". Turning
something off is a setting in its own right:

| You want | You write | Reads back as |
| --- | --- | --- |
| this colour | `setSeriesFill(c, 1, "FF0000")` | a fill with that colour |
| deliberately invisible | `setSeriesFill(c, 1, :none)` | a fill with `kind === :none` |
| let the style decide | `setSeriesFill(c, 1, :inherit)` | `nothing` |

The same three states apply to lines, markers, labels and titles.

## Formatting resolves up a cascade

Excel's chart formatting inherits: a data point falls back to its series, a
series to its group, a group to the plot area and the chart space. The resolvers
walk that chain and return an [`Effective`](@ref XLSX.Charts.Effective), which
carries the value, the rung that supplied it, and the whole chain:

```julia
julia> e = getSeriesFill(c, 1);

julia> e.value.fgcolor.rgb
"FF0000"

julia> e.site.level
:series
```

A `value` of `nothing` means no rung wrote the property, so Excel takes it from
the chart's style part. [Formatting charts](formattingCharts.md) covers this in
full.

## Created charts follow the workbook's theme

Charts made by [`addChart`](@ref) and [`addChartEx`](@ref) start from Excel's own
default chart parts, with the colours written as theme references. So a new
chart looks like one Excel would have made in that workbook — which means a
chart created in a workbook from `XLSX.newxlsx()` shows that template's older
Office colours rather than current Excel's.