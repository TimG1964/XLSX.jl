```@meta
CurrentModule = XLSX.Charts
```

!!! note "Experimental"

    Handling of native Excel charts in XLSX.jl is experimental and a work in 
    progress. All aspects may be subject to change. Feedback is welcome!


# Formatting charts

```julia
julia> using XLSX, XLSX.Charts, Colors
```

Chart appearance is DrawingML: the same fills, outlines and text properties that
describe any shape in an Office document. This page covers reading them,
changing them, and the inheritance that decides what a chart actually looks
like.

## Fills and outlines

Every drawable part of a chart carries shape properties (`c:spPr`): the series
graphic, individual data points, the plot area, the chart area, axes, the
legend, gridlines, trendlines and error bars. The getters return a
[`DrawingShapeProps`](@ref), whose `fill` and `line` are a
[`DrawingFill`](@ref) and a [`DrawingLine`](@ref):

```julia
julia> f = XLSX.opentemplate("chart_appearance.xlsx");

julia> c = getCharts(f)[1];

julia> sp = getSeriesShapeProps(c, 1)
DrawingShapeProps(fill solid, no line, effects)

julia> sp.fill
XLSX.Charts.DrawingFill solid
  foreground: #156082

julia> sp.line.fill.kind          # an `a:ln` is present, holding <a:noFill/>
:none
```

Setting them is by property rather than by struct:

```julia
julia> setSeriesFill(c, 1, "FF0000");

julia> setSeriesLine(c, 1; color = "000000", width = 1.5, dash = :sysDot);
```

[`setSeriesLine`](@ref) applies every keyword in one rebuild of the chart part,
which is both faster and safer than a sequence of single-property calls: if one
value is rejected, nothing is written. The single-property setters —
[`setSeriesLineColor`](@ref), [`setSeriesLineWidth`](@ref),
[`setSeriesLineDash`](@ref), [`setSeriesLineCap`](@ref),
[`setSeriesLineCompound`](@ref), [`setSeriesLineJoin`](@ref) and
[`setSeriesLineMiterLimit`](@ref) — exist for when you want just one.

Widths are points, as in Excel's own boxes. Dash, cap and compound take the
DrawingML name or Excel's name for the same thing, whichever you find easier to
remember:

| | DrawingML | Excel |
| --- | --- | --- |
| dash | `:sysDot`, `:sysDash`, `:lgDash`, `:lgDashDot`, `:lgDashDotDot` | `:roundDot`, `:squareDot`, `:longDash`, `:longDashDot`, `:longDashDotDot` |
| cap | `:sq`, `:rnd` | `:square`, `:round` |
| compound | `:sng`, `:dbl`, `:tri` | `:simple`, `:double`, `:triple` |

`:solid`, `:dash` and `:dashDot`; `:flat`; and `:thickThin` and `:thinThick` are
spelled the same either way.

Setters take symbols; getters report strings, because that is what the file
holds. So `dash = :sysDot` reads back as `"sysDot"`, and an Excel name you set
comes back as its DrawingML equivalent — `dash = :longDash` reads back as
`"lgDash"`.

## Three states, not two

`nothing`, `:none` and `:inherit` mean three different things, and the
difference survives a round trip:

| Written | Reads back as | Excel draws |
| --- | --- | --- |
| nothing at all | `nothing` | whatever the chart style says |
| `<a:noFill/>` | a fill with `kind === :none` | nothing |
| `<a:solidFill>` | a fill with `kind === :solid` | that colour |

So `setSeriesFill(c, 1, :none)` makes a series deliberately invisible, while
`setSeriesFill(c, 1, :inherit)` removes the setting and lets the style decide.
The same three states apply to lines, markers, labels and titles.

[`has_fill`](@ref) and [`has_line`](@ref) collapse those three cases into the
one question a call site usually asks — does this element say something is
drawn?

```julia
julia> sp = getSeriesShapeProps(c, 1);

julia> has_fill(sp), has_line(sp)
(true, true)
```

They take a `DrawingShapeProps`, a fill or line on its own, or the
[`Effective`](@ref) a resolver returns. `false` never means Excel draws nothing:
a property written at no rung comes from the chart style, which this layer does
not read.

## Colours

Anywhere a colour is set, four forms are accepted:

```julia
julia> setSeriesFill(c, 1, "FF0000");                  # hex RGB or AARRGGBB

julia> setSeriesFill(c, 1, "coral");                   # any Colors.jl named colour

julia> setSeriesFill(c, 1, colorant"seagreen");        # a Colors.jl value

julia> setSeriesFill(c, 1, SchemeColor(:accent2));     # a theme colour
```

A [`SchemeColor`](@ref) names a slot in the workbook's theme rather than a fixed
colour, so it follows the theme the way Excel's own formatting does. Transforms
are percentages, applied in the order given:

```julia
julia> SchemeColor(:tx1; lumMod = 65, lumOff = 35)
SchemeColor(:tx1, [:lumMod => 65.0, :lumOff => 35.0])
```

A colour read back from a file is a [`DrawingColor`](@ref), which records the
reference as authored *and* the RGB it resolves to:

```julia
julia> col = getSeriesFill(c, 1).value.fgcolor;

julia> col.kind, col.val, col.rgb
(:scheme, "accent2", "E97132")
```

So a theme colour stays a theme colour in the file, while `rgb` tells you what
it looks like.

## The cascade

Excel's chart formatting inherits. A data point falls back to its series, a
series to its group, a group to the plot area and then the chart space. The
resolvers walk that chain and return an [`Effective`](@ref):

```julia
julia> setSeriesFill(c, 1, "FF0000");

julia> e = getSeriesFill(c, 1);

julia> e.value.fgcolor.rgb      # the colour now in force
"FF0000"

julia> e.site.level             # which rung supplied it
:series
```

`site` is the rung that answered; `level` names it — `:point`, `:series`,
`:group`, `:plotarea` or `:chartspace`. `chain` holds the rungs that exist in
the file, highest precedence first, which is how a setter knows where to write.
A data point appears in the chain only once the file has a `c:dPt` for it, so a
series whose points carry no overrides resolves through a chain of one.

`value === nothing` means no rung wrote the property at all, so Excel takes it
from the chart's style part — not that the property is off. `chain` is populated
either way.

To resolve a point rather than a series, pass `point`. The point's own
formatting wins where it has any, and falls back to the series where it hasn't:

```julia
julia> getSeriesFill(c, 1; point = 2).site.level     # this point has a c:dPt
:point

julia> getSeriesFill(c, 1; point = 3).site.level     # this one doesn't
:series
```

`point` counts from 1, while a [`ChartDataPoint`](@ref)'s `idx` is the `c:idx`
Excel wrote, counting from 0 — so a point read back at `idx == 1` is `point = 2`
here.

!!! note
    The cascade stops at the chart XML. Excel's style and colour parts
    (`style1.xml`, `colors1.xml`) are not read, so a property written at no rung
    resolves to `nothing` rather than to the style's value. See
    [Limitations](chartLimitations.md).

## Markers

Markers belong to line, scatter and radar series, and to individual points of
them. Setting one on a bar or pie series throws, because `c:barSer` has no
`c:marker` child in the schema — which is why this section uses a line chart.

```julia
julia> k = XLSX.opentemplate("chart_kinds.xlsx");

julia> lc = getCharts(k["linemarkers"])[1];

julia> setMarker(lc, 1; symbol = :circle, size = 7, fill = "FF0000");

julia> getSeriesMarker(lc, 1)
XLSX.Charts.ChartMarker (series idx 0)
  symbol: circle
  size: 7 pt
  shape: DrawingShapeProps(fill solid, line 0.75pt, effects)

julia> setMarkerSymbol(lc, 1, :diamond; point = 3);     # just this point

julia> getDataPointMarker(lc, only(getSeriesDataPoints(lc, 1))).symbol
:diamond
```

`point` counts from 1. A point's marker overrides the series' marker for that
point; [`getMarkerFill`](@ref) resolves the two rungs the same way
[`getSeriesFill`](@ref) resolves shape properties.

The symbols are `:circle`, `:dash`, `:diamond`, `:dot`, `:none`, `:plus`,
`:square`, `:star`, `:triangle`, `:x`, plus `:auto` to let Excel choose and
`:picture` for a marker drawn from an image. Excel uses the same words, so
there are no aliases.

A marker whose `symbol` is `:none` is explicitly turned off, which is not the
same as a series with no `c:marker` element at all.

## Text

Text formatting is set field by field, because text inherits field by field — a
size from one rung and a weight from another both apply:

```julia
julia> setChartTitleTextProp(c, :size, 18);

julia> setChartTitleTextProp(c, :bold, true);

julia> setLabelTextProp(c, 1, :italic, true);          # a series' data labels
```

The fields are those of [`DrawingRunProps`](@ref): `:size`, `:bold`, `:italic`,
`:under`, `:strike`, `:caps`, `:baseline`, `:kern`, `:spacing`, `:latin`, `:ea`,
`:cs`, `:lang`, plus `:fill` and `:line` for the text's own colour and outline.

[`getLabelTextProp`](@ref) is the resolving counterpart. It takes a field name
and walks the chain from the series up through the group, plot area and chart
space, returning an [`Effective`](@ref) like the shape resolvers:

```julia
julia> getLabelTextProp(c, 1, :size).value             # the title's 18 doesn't apply
10.5
```

`:fill` and `:line` are composite: they resolve to a whole
[`DrawingFill`](@ref) or [`DrawingLine`](@ref), since a run's fill is written at
some rung or not at all — so `getLabelTextProp(c, 1, :fill).value.fgcolor` is a
field access on the result, not a further walk.

To read what one element says rather than what applies, the `…TextProps` getters
— [`getChartTitleTextProps`](@ref), [`getAxisTextProps`](@ref),
[`getLegendTextProps`](@ref), [`getDataLabelTextProps`](@ref) and the rest —
return a [`DrawingText`](@ref): a text body of paragraphs and runs, each
carrying its own `DrawingRunProps`. The [Chart types](../api/chartTypes.md) page
documents the structure.

## Titles and legends

A title's text and its formatting are separate. [`getChartTitleText`](@ref)
gives the literal rich text, [`getChartTitleRef`](@ref) the formula when the
title is bound to a cell, and [`getChartTitleTextProps`](@ref) the formatting
that applies either way.

```julia
julia> setChartTitleText(c, "Revenue by region");

julia> setAxisTitleText(c, getChartAxes(c, :value)[1], "£m");
```

Multi-line titles are written as a [`DrawingText`](@ref) with several
paragraphs:

```julia
julia> setChartTitleText(c, XLSX.Charts.DrawingText("Revenue", "by region"));

julia> getChartTitle(c)
"Revenue\nby region"
```

A `"\n"` inside a single string gives the same two-line title.

Legends are read through [`getChartLegend`](@ref) and its companions —
[`getLegendPos`](@ref), [`getLegendOverlay`](@ref),
[`getLegendShapeProps`](@ref), [`getLegendTextProps`](@ref) — and their text is
set with [`setLegendTextProp`](@ref).

## Axes and gridlines

Axis appearance splits the same way: [`getAxisShapeProps`](@ref) for the axis
line, [`getAxisTextProps`](@ref) for the tick labels,
[`getAxisGridlines`](@ref) for the gridlines:

```julia
julia> ax = getChartAxes(c, :value)[1];

julia> g = getAxisGridlines(c, ax);

julia> isnothing(g)          # are there gridlines at all?
false
```

Gridlines are one of several features where presence is the state: `nothing`
means the element is absent and none are drawn; an element present with no
`c:spPr` means they are drawn with inherited formatting, and comes back as a
`DrawingShapeProps` with every field absent. The same pattern covers drop lines
([`getGroupDropLines`](@ref)), high-low lines ([`getGroupHiLowLines`](@ref)),
series lines ([`getGroupSeriesLines`](@ref)) and up-down bars
([`getGroupUpDownBars`](@ref)).

The axis scale itself — bounds, units, tick marks, label position, number format
— is read through the `getAxis…` accessors listed on the
[Charts](../api/charts.md) API page. `nothing` there means Excel chooses
automatically, which is the usual case.
