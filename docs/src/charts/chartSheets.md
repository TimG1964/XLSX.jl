```@meta
CurrentModule = XLSX.Charts
```

!!! note "Experimental"

    Handling of native Excel charts in XLSX.jl is experimental and a work in 
    progress. All aspects may be subject to change. Feedback is welcome!


# Sheets with charts

```julia
julia> using XLSX, XLSX.Charts
```

Copying, renaming and deleting sheets are core operations, but a sheet holding a
chart carries more than cells: a drawing part, one chart part per chart, and the
style and colour parts beside it. This page is what happens to those.

## Copying a sheet

[`XLSX.copysheet!`](@ref) clones the chart parts along with the cells, and
repoints the copies at the new sheet:

```julia
julia> f = XLSX.opentemplate("chart_basic.xlsx");

julia> XLSX.copysheet!(f["Data"], "Copy");

julia> [(c.name, c.sheet) for c in getCharts(f)]
2-element Vector{Tuple{String, String}}:
 ("chart1", "Data")
 ("chart2", "Copy")
```

Each copy is a part of its own, so formatting one afterwards leaves the other
alone. Both schemas are cloned the same way, so a sheet holding a `c:` chart and
a `cx:` chart gives a copy of each.

Series references follow the copy: a chart on `Copy` plots from `Copy`, not from
the sheet it was copied from.

## Renaming a sheet

[`XLSX.renamesheet!`](@ref) repoints the series references of every chart that
plots from the renamed sheet, wherever the chart itself lives:

```julia
julia> XLSX.renamesheet!(f["Data"], "Sales");

julia> getChartRanges(getCharts(f)[1])[1].values
Sales!B2:B5
```

Chart parts store sheet-qualified references, so this matters: a rename that
missed them would leave the chart plotting from a sheet that no longer exists.
Cached values are untouched, since the data hasn't changed.

## Deleting a sheet

[`XLSX.deletesheet!`](@ref) removes the charts anchored to the sheet, along with
their drawing, style and colour parts and the content-type entries for all of
them:

```julia
julia> g = XLSX.opentemplate("chart_basic.xlsx");

julia> XLSX.addsheet!(g, "Keep");          # a workbook needs at least one sheet

julia> XLSX.deletesheet!(g, "Data");

julia> isempty(getCharts(g))
true
```

A chart elsewhere in the workbook that plots from the deleted sheet stays, but
loses the reference, as it does in Excel. A `c:` chart's reference becomes
`#REF!` and its cached values are kept, so the chart still shows what it last
displayed; [`getChartRanges`](@ref) reports the broken reference as
`XLSX.XL_REF`:

```julia
julia> h = XLSX.opentemplate("Chartsheet.xlsx");

julia> XLSX.addsheet!(h, "Keep");

julia> XLSX.deletesheet!(h, "Sheet1");

julia> getChartRanges(getCharts(h)[1])[1].values
XL_REF::CellErrorType = 0x0000000000000004
```

A `cx:` chart loses the reference outright, as Excel does it: the data is
emptied, and a series name or title linked to a cell on the sheet is removed, so
its ranges read as `nothing`.

## Chartsheets

A chartsheet is a sheet in its own right, so it is deleted like any other:

```julia
julia> addChart(g, :line; sheetname = "Trend");

julia> XLSX.deletesheet!(g, "Trend");
```

Deleting the worksheet a chartsheet's chart plots from leaves the chartsheet in
place, with its references invalidated as described above.

Copying a chartsheet is not supported: [`XLSX.copysheet!`](@ref) throws an
`XLSXError`. Create a second chart with [`addChart`](@ref) instead.

## What a sheet's charts cost

Deleting or copying a sheet touches more parts than the sheet's own XML, so a
workbook written after either is not byte-identical to one that never had the
chart. That is expected: the package rewrites the package-level bookkeeping —
relationships, content types, part names — rather than patching it, and Excel
reads the result as its own.