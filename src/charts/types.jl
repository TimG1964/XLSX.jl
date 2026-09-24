# ===========================================================================
# Charts
# ===========================================================================

"""
    ChartRef

One cached reference from a chart series: the formula it came from, the number
format Excel recorded for it, and the cached values themselves.

# Fields
- `kind::Symbol` - one of `:num`, `:str`, `:multiLvlStr`, `:numLit`, `:strLit`.
- `ref::Union{Nothing,String}` - the `c:f` formula (`Sheet1!\$B\$2:\$B\$9`).
  `nothing` for literal (`c:numLit` / `c:strLit`) series, which have no source range.
- `format_code::Union{Nothing,String}` - number format recorded in the cache.
- `ptCount::Int` - number of points Excel declared, whether or not the cache was read.
- `data::Vector` - cached values, length `ptCount`, gaps as `missing`. Empty when
  the chart was read with `cache=false`.
- `errors::Dict{Int,UInt64}` - index => error code for cached error values.

Excel caches only `#N/A` as an error. Other error values are cached as 0, so they
cannot be told apart from a real zero.

`ChartRef` only applies to `c:` charts and not `cx:` charts.


!!! note
    For `kind == :multiLvlStr` each element of `data` is itself a level vector,
    in document order (Excel writes the innermost/leaf level first).
"""
struct ChartRef
    kind::Symbol
    ref::Union{Nothing,String}
    format_code::Union{Nothing,String}
    ptCount::Int
    data::Vector
    errors::Dict{Int,UInt64}
end

# chart part path => (sheet name, from, to, rId)
const ChartAnchor = NamedTuple{
    (:sheet, :from, :to, :rId),
    Tuple{String,Union{Nothing,String},Union{Nothing,String},String},
}

# Internal: one discovered chart part, before parsing.
struct ChartLocation
    path::String
    anchor::Union{Nothing,ChartAnchor}
    schema::Symbol   # :c or :cx
end

"""
    ChartSeries

One series (`c:ser`), as read.

`categories` holds `c:cat` for category charts and `c:xVal` for scatter and bubble
charts; `values` holds `c:val` or `c:yVal` correspondingly, so the two fields mean
the same thing whatever the chart type.

`idx` is the series' `c:idx`, unique within the chart and unaffected by adding
series. It is the key the value types carry as `series_idx`, so a
[`ChartDataPoint`](@ref) or [`ChartTrendline`](@ref) read from this series still
addresses it after later writes. `order` is the plotting order (`c:order`) and is
not a key.

A series position `i`, as taken by the setter functions and the setters, counts
from 1 in document order and is resolved afresh on each call. It is not `idx`;
the two agree only where Excel has numbered the series from 0 in order, as it
normally does.
"""
struct ChartSeries
    idx::Int
    order::Int
    charttype::Symbol
    name::Union{Nothing,String}
    name_ref::Union{Nothing,ChartRef}
    categories::Union{Nothing,ChartRef}
    values::Union{Nothing,ChartRef}
    bubble_sizes::Union{Nothing,ChartRef}
    raw::Union{Nothing,XML.Node}          # the c:ser element
end

"""
    ChartRange

A resolved series reference: a `SheetCellRef`, `SheetCellRange`, `SheetRowRange`,
`SheetColumnRange` or `NonContiguousRange`, or `nothing` where the series has no
reference of that kind or it could not be resolved.
"""
const ChartRange = Union{Nothing,SheetCellRef,SheetCellRange,SheetRowRange,SheetColumnRange,NonContiguousRange}

"""
    ChartRanges

What [`getChartRanges`](@ref) returns for one series: a named tuple of `idx`,
`name`, and the `categories`, `values` and `bubble_sizes` references, each a
[`ChartRange`](@ref).

`idx` is the series' `c:idx`; see [`ChartSeries`](@ref). A field is `nothing`
where the series has no reference of that kind — `bubble_sizes` on anything but
a bubble chart, for instance.
"""
const ChartRanges = @NamedTuple{
    idx::Int,
    name::Union{Nothing,String},
    categories::ChartRange,
    values::ChartRange,
    bubble_sizes::ChartRange,
}


"""
    AbstractChart

Supertype for charts read from a workbook. Two concrete subtypes exist:
[`Chart`](@ref) for the standard `c:` schema, and [`ChartEx`](@ref) for the
newer `cx:` schema used by waterfall, funnel, treemap, sunburst, histogram,
Pareto, box & whisker and region map charts.
"""
abstract type AbstractChart end

"""
    Chart

A handle to one `c:` chart part. It holds the part's identity and anchor only;
everything else is read from the current part on each call, so a `Chart` never
goes stale and setters return it unchanged.

# Fields
- `package` — the `XLSXFile` the part belongs to.
- `path` — package path, e.g. `"xl/charts/chart1.xml"`.
- `name` — part name without extension, e.g. `"chart1"`.
- `rId` — relationship id of the chart within its drawing part, if resolved.
- `sheet` — name of the sheet the chart is anchored to, if resolved.
- `from`, `to` — anchor cell references as strings, following `getImages`.

Content is reached through accessors: [`getChartTitle`](@ref),
[`getChartTypes`](@ref), [`getChartSeries`](@ref), [`getChartData`](@ref) and the
rest of the chart API. The values they return carry keys, so they remain valid
arguments after later writes; see [`ChartSeries`](@ref).
"""
struct Chart <: AbstractChart
    package::XLSXFile
    path::String
    name::String
    rId::Union{Nothing,String}
    sheet::Union{Nothing,String}
    from::Union{Nothing,String}
    to::Union{Nothing,String}
end

"""
    ChartEx

A handle to one chartEx part: a chart in the Microsoft `cx:` namespace
(`http://schemas.microsoft.com/office/drawing/2014/chartex`), used by Excel for
waterfall, funnel, treemap, sunburst, histogram, Pareto, box & whisker and
region map charts. Charts in the ECMA-376 `c:` namespace are [`Chart`](@ref).

A `ChartEx` holds only the part's identity and its anchor. Its content (layout,
data references, title) is read from the part on each call, so a `ChartEx`
remains valid after the part is written to.

# Fields
- `package` - the `XLSXFile` containing the chart.
- `path` - package path, e.g. `"xl/charts/chartEx1.xml"`.
- `name` - part name without extension, e.g. `"chartEx1"`.
- `rId` - relationship id of the chart within its drawing part, if resolved.
- `sheet` - name of the sheet the chart is anchored to, if resolved.
- `from`, `to` - anchor cell references as strings, following `getImages`.

# Reading content
- [`getChartType`](@ref) - e.g. `:waterfall`, `:histogram`.
- [`getChartTitle`](@ref) - title text, whether typed or bound to a cell;
  `nothing` if the title has no text of its own.
- [`getChartRanges`](@ref) - the source range of each data dimension, resolved
  through the workbook's hidden `_xlchart.*` defined names.

See also [`getCharts`](@ref), [`getChartSchema`](@ref).
"""
struct ChartEx <: AbstractChart
    package::XLSXFile
    path::String
    name::String
    rId::Union{Nothing,String}
    sheet::Union{Nothing,String}
    from::Union{Nothing,String}
    to::Union{Nothing,String}
end

"""
    ChartExDimension

One data dimension of a chartEx data block: a `cx:numDim` or `cx:strDim`.

# Fields
- `kind` - `:num` or `:str`.
- `type` - the dimension's role as written, e.g. `:val`, `:cat`, `:size`,
  `:x`, `:y`, `:colorVal`, `:colorStr`, `:entityId`.
- `formula` - the text of `cx:f`, usually a hidden `_xlchart.*` defined name;
  `nothing` if the dimension has no formula.
- `range` - the resolved source range; `nothing` if unresolved.
"""
struct ChartExDimension
    kind::Symbol
    type::Symbol
    formula::Union{Nothing,String}
    range::ChartRange
end

"""
    ChartExData

One `cx:data` block. Series refer to a block by `id` through `cx:dataId`.
"""
struct ChartExData
    id::Int
    dimensions::Vector{ChartExDimension}
end

"""
    ChartExBinning

Histogram binning from `cx:layoutPr/cx:binning`. Field names follow the XML.
A numeric field may also be `:auto`; `nothing` means not written.

# Fields
- `intervalClosed` - `:r` or `:l`, the closed side of each bin interval.
- `underflow`, `overflow` - the underflow and overflow bin cut-offs.
- `binSize` - bin width.
- `binCount` - number of bins.
"""
struct ChartExBinning
    intervalClosed::Union{Nothing,Symbol}
    underflow::Union{Nothing,Float64,Symbol}
    overflow::Union{Nothing,Float64,Symbol}
    binSize::Union{Nothing,Float64,Symbol}
    binCount::Union{Nothing,Int,Symbol}
end

const SCHEME_TOKENS = (:bg1, :tx1, :bg2, :tx2, :accent1, :accent2, :accent3,
                       :accent4, :accent5, :accent6, :hlink, :folHlink,
                       :lt1, :dk1, :lt2, :dk2)

const COLOR_TRANSFORMS = (:lumMod, :lumOff, :shade, :tint, :satMod, :alpha)

# bg1/lt1, tx1/dk1, bg2/lt2 and tx2/dk2 name the same slot. The field keeps the
# token as written so a round trip preserves it; equality and hashing normalize
# so two spellings of one color compare equal.
const SCHEME_ALIASES = Dict(:lt1 => :bg1, :dk1 => :tx1, :lt2 => :bg2, :dk2 => :tx2)

"""
    SchemeColor(token; lumMod = nothing, lumOff = nothing, ...)
    SchemeColor(token, transforms)

A DrawingML theme color: a token naming a slot in the workbook's color scheme,
plus any transforms applied to it. Distinct from the spreadsheet
`<color theme="N"/>` mechanism, which is index-ordered and uses a different
tint algorithm — see `get_theme_colors` for that.

`token` is one of `:bg1`, `:tx1`, `:bg2`, `:tx2`, `:accent1` … `:accent6`,
`:hlink`, `:folHlink`, or the aliases `:lt1`, `:dk1`, `:lt2`, `:dk2`.

The `:lt1`, `:dk1`, `:lt2` and `:dk2` aliases are preserved as written, so a
round trip keeps the spelling the file used, but two spellings of one slot
compare and hash equal.

Transforms are held in document order, because DrawingML applies them in
sequence and `lumMod` before `lumOff` is not the same as the reverse. Equality
is positional for the same reason: two `SchemeColor`s with the same transforms
in different orders are not equal, because they do not render the same. Values
are percentages, not the hundredths DrawingML writes: `lumMod = 75` is 75%.

The keyword constructor emits transforms in the conventional order — `lumMod`
before `lumOff`, shade or tint last. Any other order needs the vector form.

# Examples

    SchemeColor(:accent1)
    SchemeColor(:accent1; lumMod = 75)                    # 104862 against the Office theme
    SchemeColor(:tx1; lumMod = 65, lumOff = 35)           # 595959
    SchemeColor(:accent2, [:shade => 50.0, :alpha => 80.0])
"""
struct SchemeColor
    token::Symbol
    transforms::Vector{Pair{Symbol,Float64}}

    function SchemeColor(token::Symbol, transforms::Vector{Pair{Symbol,Float64}})
        token in SCHEME_TOKENS || throw(XLSXError(
            "`$token` is not a theme color token. Valid tokens: " *
            join(SCHEME_TOKENS, ", ") * "."))
        for (name, _) in transforms
            name in COLOR_TRANSFORMS || throw(XLSXError(
                "`$name` is not a color transform. Valid transforms: " *
                join(COLOR_TRANSFORMS, ", ") * "."))
        end
        return new(token, transforms)
    end
end

function SchemeColor(token::Symbol; lumMod = nothing, lumOff = nothing,
                     satMod = nothing, shade = nothing, tint = nothing,
                     alpha = nothing)
    t = Pair{Symbol,Float64}[]
    isnothing(lumMod) || push!(t, :lumMod => Float64(lumMod))
    isnothing(lumOff) || push!(t, :lumOff => Float64(lumOff))
    isnothing(satMod) || push!(t, :satMod => Float64(satMod))
    isnothing(shade)  || push!(t, :shade  => Float64(shade))
    isnothing(tint)   || push!(t, :tint   => Float64(tint))
    isnothing(alpha)  || push!(t, :alpha  => Float64(alpha))
    return SchemeColor(token, t)
end

canonical_token(t::Symbol) = get(SCHEME_ALIASES, t, t)

Base.:(==)(a::SchemeColor, b::SchemeColor) =
    canonical_token(a.token) == canonical_token(b.token) && a.transforms == b.transforms

Base.hash(c::SchemeColor, h::UInt) =
    hash(c.transforms, hash(canonical_token(c.token), hash(:SchemeColor, h)))

Base.isequal(a::SchemeColor, b::SchemeColor) = a == b

"""
    DrawingColor

A DrawingML colour: the element as written, plus the RGB it resolves to.

`kind` and `val` record the reference as authored - a theme colour stays a
theme colour - and `transforms` the modifications applied to it, in document
order. `rgb` and `alpha` give the resolved result for anyone who just wants to
know what it looks like.

# Fields
- `kind::Symbol` - `:srgb`, `:scheme`, `:sys`, `:prst`, `:hsl` or `:scrgb`.
- `val::String` - the `val` attribute: `"FF0000"`, `"accent1"`, `"windowText"`.
- `transforms::Vector{Pair{Symbol,Int}}` - e.g. `[:lumMod => 60000, :lumOff => 40000]`,
  in thousandths of a percent, in the order DrawingML applies them. A
  [`SchemeColor`](@ref) built in code holds the same transforms as percentages,
  so `60000` here is `lumMod = 60` there.
- `rgb::String` - the resolved colour as `"RRGGBB"`.
- `alpha::Float64` - `1.0` unless an `alpha` transform applies.

`rgb` and `alpha` are resolved values, not source, and are not written — the
file keeps the scheme reference and its transforms so the theme still applies.
"""
struct DrawingColor
    kind::Symbol
    val::String
    transforms::Vector{Pair{Symbol,Int}}
    rgb::String
    alpha::Float64
end

"""
    ColorSpec

Anything that can specify a color where one is written: a parsed
[`DrawingColor`](@ref), a [`SchemeColor`](@ref), or a string naming an
`AARRGGBB` value or a Colors.jl color.

The parser only ever produces `DrawingColor`, so a value read from a file is
always that. The other two exist for construction.
"""
const ColorSpec = Union{DrawingColor, SchemeColor, String}

"""
    DrawingFill

A DrawingML fill: `<a:solidFill>`, `<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`
or `<a:blipFill>`.

A solid fill resolves to one colour in `fgcolor`. A pattern fill resolves to
two, `fgcolor` and `bgcolor`, with the pattern itself in `preset` - matching
how cell fills are exposed. Gradient and picture fills are identified by `kind`
but not modelled further: `raw` holds the element as read, so nothing is lost
on write.

# Fields
- `kind::Symbol` - `:none`, `:solid`, `:gradient`, `:pattern`, `:blip` or `:group`.
- `fgcolor` - the colour of a solid fill, or a pattern's foreground.
- `bgcolor` - a pattern's background; `nothing` otherwise.
- `preset::Union{Nothing,String}` - a pattern's `prst` attribute, e.g. `"pct25"`,
  `"ltUpDiag"`.
- `raw::Union{Nothing,XML.Node}` - the element as read, or `nothing` for a fill
  built in code. Gradient, pattern, picture and group fills are written back from
  it.
"""
struct DrawingFill
    kind::Symbol
    fgcolor::Union{Nothing,ColorSpec}
    bgcolor::Union{Nothing,ColorSpec}
    preset::Union{Nothing,String}
    raw::Union{Nothing,XML.Node}
end

"""
    DrawingFill(kind; fgcolor = nothing, bgcolor = nothing, preset = nothing)

A fill built rather than parsed. `kind` is `:none`, `:solid`, `:gradient`,
`:pattern`, `:blip` or `:group`; only `:none` and `:solid` can be serialized
from a constructed value, since the others are modelled partially and written
back from `raw`.
A fill read from a file carries a DrawingColor with `rgb` resolved; one built
by hand may carry a SchemeColor or a string, which has no resolved value until
it is written and read back.
"""
DrawingFill(kind::Symbol; fgcolor = nothing, bgcolor = nothing, preset = nothing) =
    DrawingFill(kind, fgcolor, bgcolor, preset, nothing)

"""
    DrawingLine

A DrawingML outline: `<a:ln>`.

The stroke colour lives in `fill`, since a line is filled the same way a shape
is - solid, gradient, pattern or none.

# Fields
- `fill::Union{Nothing,DrawingFill}` - the stroke; `nothing` where the element
  says nothing about it, `kind === :none` where it explicitly has no outline.
- `width::Union{Nothing,Float64}` - the `w` attribute, in points (the file stores
  EMU, 12700 to the point).
- `dash::Union{Nothing,String}` - `"solid"`, `"dash"`, `"sysDot"`, and so on.
- `cap::Union{Nothing,String}` - `"rnd"`, `"sq"`, `"flat"`.
- `compound::Union{Nothing,String}` - the `cmpd` attribute: `"sng"`, `"dbl"`, …
- `join::Union{Nothing,String}` - `"round"`, `"bevel"` or `"miter"`, from the
  `a:round`/`a:bevel`/`a:miter` child.
- `miter_limit::Union{Nothing,Float64}` - the `a:miter` `lim`, as a multiple of
  the line width; only with a miter join.
- `raw::Union{Nothing,XML.Node}` - the element as read, or `nothing` for a line
  built in code.
"""
struct DrawingLine
    fill::Union{Nothing,DrawingFill}
    width::Union{Nothing,Float64}         # w, points (file stores EMU)
    dash::Union{Nothing,String}
    cap::Union{Nothing,String}
    compound::Union{Nothing,String}
    join::Union{Nothing,String}           # :round, :bevel, :miter
    miter_limit::Union{Nothing,Float64}   # multiple of line width; only with miter
    raw::Union{Nothing,XML.Node}
end

DrawingLine(; fill = nothing, width = nothing, dash = nothing, cap = nothing,
              compound = nothing, join = nothing, miter_limit = nothing) =
    DrawingLine(fill, width, dash, cap, compound, join, miter_limit, nothing)

"""
    DrawingShapeProps

Shape properties (`a:spPr`) — the fill and outline of anything drawn in a chart:
series, data points, the plot area, the chart area, axis lines, legend, gridlines.

`fill` and `line` distinguish three states, and the difference matters:

| file                            | reads as               | means                  |
|---------------------------------|------------------------|------------------------|
| no fill element                 | `nothing`              | inherit from the style |
| `<a:noFill/>`                   | `kind == :none`        | deliberately invisible |
| `<a:solidFill>…`                | `kind == :solid`       | this colour            |

The same applies to `line.fill`: `<a:ln><a:noFill/></a:ln>` is how Excel writes
"no border", which is not the same as omitting `a:ln` entirely.

`effects` holds `a:effectLst` or `a:effectDag` as an unparsed node — enough to
report that a shape has effects without modelling shadows and glows. Geometry 
(`a:xfrm`, `a:prstGeom`, `a:custGeom`) and 3-D (`a:scene3d`, `a:sp3d`) are
not modelled; they stay in `raw` and are written back from it unchanged. Chart parts
rarely carry geometry — it belongs to the drawing shapes that host the chart, not the
chart itself.
"""
struct DrawingShapeProps
    fill::Union{Nothing,DrawingFill}
    line::Union{Nothing,DrawingLine}
    effects::Union{Nothing,XML.Node}    # a:effectLst or a:effectDag
    bwmode::Union{Nothing,String}       # bwMode: clr | auto | gray | ltGray | invGray | ...
    raw::Union{Nothing,XML.Node}
end

# =============================================================================
# Every optional field is Union{Nothing,T}: absent means "inherit from the
# theme or the parent list style", which is not the same as an explicit value,
# and the difference has to survive a round trip.
# =============================================================================

"""
    DrawingRunProps

Character-level properties: `a:rPr`, `a:defRPr` or `a:endParaRPr`.

Sizes are points (the file stores 1/100 pt), `baseline` is a fraction of the font
size, and `under` / `strike` / `caps` keep the DrawingML vocabulary as written
(`"sng"`, `"noStrike"`, `"small"`). Typefaces may be theme references —
`"+mn-lt"` for the minor latin font, `"+mj-lt"` for major.
"""
struct DrawingRunProps
    lang::Union{Nothing,String}
    size::Union{Nothing,Float64}        # sz, points
    bold::Union{Nothing,Bool}           # b
    italic::Union{Nothing,Bool}         # i
    under::Union{Nothing,String}        # u
    strike::Union{Nothing,String}
    caps::Union{Nothing,String}         # cap
    baseline::Union{Nothing,Float64}    # fraction of the font size
    kern::Union{Nothing,Float64}        # points
    spacing::Union{Nothing,Float64}     # spc, points
    fill::Union{Nothing,DrawingFill}
    line::Union{Nothing,DrawingLine}    # a:ln — text outline
    latin::Union{Nothing,String}
    ea::Union{Nothing,String}
    cs::Union{Nothing,String}
    raw::Union{Nothing,XML.Node}
end

DrawingRunProps(; lang = nothing, size = nothing, bold = nothing, italic = nothing,
                  under = nothing, strike = nothing, caps = nothing,
                  baseline = nothing, kern = nothing, spacing = nothing,
                  fill = nothing, line = nothing,
                  latin = nothing, ea = nothing, cs = nothing) =
    DrawingRunProps(lang, size, bold, italic, under, strike, caps, baseline,
                    kern, spacing, fill, line, latin, ea, cs, nothing)

"""
    DrawingParaProps

Paragraph properties (`a:pPr`). Margins and indent are points; spacing fields
are `(:pct, percent)` or `(:pts, points)`. `defprops` is the nested `a:defRPr`,
which in a chart `txPr` is usually the only place the font is specified.
"""
struct DrawingParaProps
    align::Union{Nothing,String}        # algn
    level::Union{Nothing,Int}           # lvl
    marginleft::Union{Nothing,Float64}  # marL, points
    marginright::Union{Nothing,Float64} # marR, points
    indent::Union{Nothing,Float64}      # points
    rtl::Union{Nothing,Bool}
    linespacing::Union{Nothing,Tuple{Symbol,Float64}}   # lnSpc
    spacebefore::Union{Nothing,Tuple{Symbol,Float64}}   # spcBef
    spaceafter::Union{Nothing,Tuple{Symbol,Float64}}    # spcAft
    defprops::Union{Nothing,DrawingRunProps}
    raw::Union{Nothing,XML.Node}
end

DrawingParaProps(; align = nothing, level = nothing,
                   marginleft = nothing, marginright = nothing, indent = nothing,
                   rtl = nothing, linespacing = nothing,
                   spacebefore = nothing, spaceafter = nothing,
                   defprops = nothing) =
    DrawingParaProps(align, level, marginleft, marginright, indent, rtl,
                     linespacing, spacebefore, spaceafter, defprops, nothing)


"""
    DrawingRun

One `a:r`, `a:br` or `a:fld`, distinguished by `kind` (`:run`, `:br`, `:fld`).
A break carries `"\\n"` as its text so `text_content` needs no special case.
"""
struct DrawingRun
    kind::Symbol
    text::Union{Nothing,String}
    props::Union{Nothing,DrawingRunProps}
    raw::Union{Nothing,XML.Node}
end

"""
    DrawingRun(text; props = nothing, kind = :run)

One run of text. `kind` is `:run`, `:br` or `:fld`; a break carries `"\\n"` as
its text.
"""
DrawingRun(text::AbstractString; props = nothing, kind::Symbol = :run) =
    DrawingRun(kind, String(text), props, nothing)


"""
    DrawingParagraph

One `a:p`: properties, runs in document order, and the trailing
`a:endParaRPr`, which is what Excel writes when a paragraph has no runs.
"""
struct DrawingParagraph
    props::Union{Nothing,DrawingParaProps}
    runs::Vector{DrawingRun}
    endprops::Union{Nothing,DrawingRunProps}
    raw::Union{Nothing,XML.Node}
end

"""
    DrawingParagraph(runs...; props = nothing, endprops = nothing)

One paragraph. Runs may be `DrawingRun`s or plain strings, which become runs
with no properties of their own — they inherit the paragraph default.
"""
DrawingParagraph(runs::Union{DrawingRun,AbstractString}...;
                 props = nothing, endprops = nothing) =
    DrawingParagraph(props,
                     DrawingRun[r isa DrawingRun ? r : DrawingRun(r) for r in runs],
                     endprops, nothing)

"""
    DrawingBodyProps

Text-body properties (`a:bodyPr`). `rotation` is degrees (stored as 1/60000),
insets are points, and `autofit` is `:none`, `:normal` or `:shape` — parsed
from a child element, with `fontscale` and `linespacereduction` populated only
for `:normal`.
"""
struct DrawingBodyProps
    rotation::Union{Nothing,Float64}           # rot, degrees
    vertical::Union{Nothing,String}            # vert
    wrap::Union{Nothing,String}
    anchor::Union{Nothing,String}
    anchorctr::Union{Nothing,Bool}
    upright::Union{Nothing,Bool}
    spcfirstlastpara::Union{Nothing,Bool}
    vertoverflow::Union{Nothing,String}
    horzoverflow::Union{Nothing,String}
    insetleft::Union{Nothing,Float64}
    insettop::Union{Nothing,Float64}
    insetright::Union{Nothing,Float64}
    insetbottom::Union{Nothing,Float64}
    autofit::Union{Nothing,Symbol}
    fontscale::Union{Nothing,Float64}          # fraction of the font size
    linespacereduction::Union{Nothing,Float64} # fraction of the font size
    raw::Union{Nothing,XML.Node}
end

DrawingBodyProps(; rotation = nothing, vertical = nothing, wrap = nothing,
                   anchor = nothing, anchorctr = nothing, upright = nothing,
                   spcfirstlastpara = nothing, vertoverflow = nothing,
                   horzoverflow = nothing,
                   insetleft = nothing, insettop = nothing,
                   insetright = nothing, insetbottom = nothing,
                   autofit = nothing, fontscale = nothing,
                   linespacereduction = nothing) =
    DrawingBodyProps(rotation, vertical, wrap, anchor, anchorctr, upright,
                     spcfirstlastpara, vertoverflow, horzoverflow,
                     insetleft, insettop, insetright, insetbottom,
                     autofit, fontscale, linespacereduction, nothing)

"""
    DrawingText

A DrawingML text body (`CT_TextBody`): `c:txPr`, `c:rich`, or the cx: equivalent.

`liststyle` is kept as an unparsed node — it is empty in most chart parts and
carries list-level defaults we don't model. `raw` is the text body element
itself, for surgical write-back.
"""
struct DrawingText
    body::Union{Nothing,DrawingBodyProps}
    liststyle::Union{Nothing,XML.Node}
    paragraphs::Vector{DrawingParagraph}
    raw::Union{Nothing,XML.Node}
end

_as_paragraphs(p::DrawingParagraph) = (p,)
_as_paragraphs(s::AbstractString) = (DrawingParagraph(line) for line in split(s, "\n"))

"""
    DrawingText(paragraphs...; body = nothing, liststyle = nothing)

A text body. Paragraphs may be `DrawingParagraph`s or plain strings, each
becoming a one-run paragraph. A string containing `"\\n"` becomes one paragraph
per line, so the two forms below are the same text body:

    DrawingText("Revenue", "by region")
    DrawingText("Revenue\\nby region")

    DrawingText(DrawingParagraph("Revenue", DrawingRun(" 2026",
                    props = DrawingRunProps(bold = true))))
"""
DrawingText(paras::Union{DrawingParagraph,AbstractString}...;
            body = nothing, liststyle = nothing) =
    DrawingText(body, liststyle,
                DrawingParagraph[q for p in paras for q in _as_paragraphs(p)],
                nothing)

"""
    ChartMarker

Marker properties (`c:marker` under a `c:ser` or `c:dPt`), as read.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `point_idx` — `c:idx` of the owning data point, 0-based as written, or
  `nothing` for the series' own marker.
- `symbol` — `:circle`, `:square`, `:none`, …; `nothing` means absent.
- `size` — points, 2 to 72; `nothing` means absent.
- `shape` — the marker's `c:spPr`, if written.
- `raw` — the element as read, or `nothing` for a marker built in code.

Identified by `series_idx` and `point_idx`. `raw` is never used to address the part.
"""
struct ChartMarker
    series_idx::Int
    point_idx::Union{Nothing,Int}        # nothing for the series' own marker
    symbol::Union{Nothing,Symbol}
    size::Union{Nothing,Int}
    shape::Union{Nothing,DrawingShapeProps}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartAxis

One axis (`c:catAx`, `c:valAx`, `c:dateAx` or `c:serAx`), as read.

# Fields
- `kind` — the element tag as a Symbol, e.g. `:valAx`.
- `axid` — `c:axId`, the key.
- `pos` — `c:axPos`: `:b`, `:t`, `:l` or `:r`.
- `crossax` — `c:crossAx`, the `axid` of the axis this one crosses.
- `deleted` — `c:delete`; a deleted axis is not drawn but remains formattable.
- `raw` — the element as read, or `nothing` for an axis built in code.

Identified by `axid`. Functions taking `(c, ax)` find the axis by it in the
current part, so `ax` stays usable across writes. The other fields describe the
axis as it was when read; getters that depend on the axis kind check the current
element, since Excel keeps the `axId` when a category axis is switched to a date
axis. `raw` is never used to address the part.
"""
struct ChartAxis
    kind::Symbol
    axid::Union{Nothing,Int}
    pos::Union{Nothing,Symbol}
    crossax::Union{Nothing,Int}
    deleted::Bool
    raw::Union{Nothing,XML.Node}
end

"""
    ChartGroup

One chart-type group in `c:plotArea` (`c:barChart`, `c:lineChart`, …), as read.

# Fields
- `kind` — the element tag as a Symbol, e.g. `:barChart`.
- `axids` — the `c:axId` values the group plots against, in order; empty for pie
  and doughnut groups.
- `raw` — the element as read, or `nothing` for a group built in code.

Identified by `(kind, axids)`: a combo chart's groups plot against different axis
pairs, so the pair is unique in practice. Functions taking `(c, g)` find the
group by it in the current part and throw if no group, or more than one, matches.
Equality and hashing use the key alone. `raw` is never used to address the part.
"""
struct ChartGroup
    kind::Symbol
    axids::Vector{Int}
    raw::Union{Nothing,XML.Node}
end

# A group's identity is its key, so equality ignores `raw`.
Base.:(==)(a::ChartGroup, b::ChartGroup) = a.kind == b.kind && a.axids == b.axids
Base.hash(g::ChartGroup, h::UInt) = hash(g.axids, hash(g.kind, hash(:ChartGroup, h)))

"""
    ChartDataPoint

Per-point formatting override on a series (`c:dPt`), as read.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `idx` — the point's `c:idx`, 0-based as written. Functions taking a point
  position count from 1, so a point at `idx == 0` is position 1.
- `invert_if_negative`, `bubble3d` — as written; `nothing` means absent.
- `raw` — the element as read, or `nothing` for a point built in code.

Identified by `series_idx` and `idx`. Functions taking `(c, d)` find the element
by these in the current part, so `d` stays usable across writes. The other fields
describe the point as it was when read; `raw` is never used to address the part.
"""
struct ChartDataPoint
    series_idx::Int                      # c:idx of the owning c:ser
    idx::Int                             # c:idx of the point, 0-based as written
    invert_if_negative::Union{Nothing,Bool}
    bubble3d::Union{Nothing,Bool}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartDataLabel

An individual data label override (`c:dLbl` within a series' `c:dLbls`), as read.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `idx` — the labelled point's `c:idx`, 0-based as written, one less than the
  position the setters take, as [`ChartDataPoint`](@ref).
- `delete` — `c:delete`; a deleted label carries no other properties.
- `raw` — the element as read, or `nothing` for a label built in code.

Identified by `series_idx` and `idx`, as [`ChartDataPoint`](@ref).
"""
struct ChartDataLabel
    series_idx::Int
    idx::Int
    delete::Union{Nothing,Bool}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartTrendline

A trendline on a series (`c:trendline`), as read.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `ordinal` — 1-based position among the series' `c:trendline` elements.
- `kind` — `c:trendlineType`: `:linear`, `:exp`, `:log`, `:movingAvg`, `:poly`, `:power`.
- `name`, `order`, `period`, `forward`, `backward`, `intercept`, `disp_rsqr`,
  `disp_eq` — as written; `nothing` means absent.
- `raw` — the element as read, or `nothing` for a trendline built in code.

Identified by `series_idx` and `ordinal`. Trendlines have no identifier of their
own, so the ordinal is positional: it holds as long as no earlier trendline on
the same series is removed. `raw` is never used to address the part.
"""
struct ChartTrendline
    series_idx::Int
    ordinal::Int                         # 1-based among the series' c:trendline elements
    kind::Union{Nothing,Symbol}
    name::Union{Nothing,String}
    order::Union{Nothing,Int}
    period::Union{Nothing,Int}
    forward::Union{Nothing,Float64}
    backward::Union{Nothing,Float64}
    intercept::Union{Nothing,Float64}
    disp_rsqr::Union{Nothing,Bool}
    disp_eq::Union{Nothing,Bool}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartErrorBars

A set of error bars on a series (`c:errBars`), as read. A series carries at most
two, one per direction.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `ordinal` — 1-based position among the series' `c:errBars` elements.
- `direction`, `bar_type`, `value_type`, `value`, `no_end_cap` — as written;
  `nothing` means absent.
- `raw` — the element as read, or `nothing` for error bars built in code.

Identified by `series_idx` and `ordinal`, as [`ChartTrendline`](@ref).
"""
struct ChartErrorBars
    series_idx::Int
    ordinal::Int                         # 1-based among the series' c:errBars elements
    direction::Union{Nothing,Symbol}
    bar_type::Union{Nothing,Symbol}
    value_type::Union{Nothing,Symbol}
    value::Union{Nothing,Float64}
    no_end_cap::Union{Nothing,Bool}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartUpDownBars

Up-down bars on a line or stock chart group (`c:upDownBars`), as read.

# Fields
- `group` — the owning [`ChartGroup`](@ref), which is the key.
- `gap_width` — `c:gapWidth`; `nothing` means absent.
- `raw` — the element as read, or `nothing` for bars built in code.
"""
struct ChartUpDownBars
    group::ChartGroup
    gap_width::Union{Nothing,Int}
    raw::Union{Nothing,XML.Node}
end

"""
    FormatSite

One rung of a formatting cascade: the node that could carry a property, and
whether it does. `container` always exists — it is the `c:ser`, `c:dPt`,
`c:marker` or chart-space element being inspected. `props` is the `spPr` or
`txPr` found on it, or `nothing` where none was written, which is the rung
being absent rather than the property being off.

`kind` distinguishes what `props` holds, so a chain is interpretable without
knowing which resolver built it: `:shape` for an `spPr` on the element itself,
`:marker` for an `spPr` on its `c:marker`, `:text` for a `txPr`.
"""
struct FormatSite
    level::Symbol                      # :point, :series, :group, :plotarea, :chartspace, :style, :theme
    kind::Symbol                       # :shape, :marker, :text
    container::XML.Node
    props::Union{Nothing,XML.Node}
end

"""
    Effective{T}

The result of resolving a property up a cascade. `value` is the first explicit
setting found, or `nothing` where the property is written at no rung at all —
which means Excel takes it from the chart style part, not that it is off. An
explicit `<a:noFill/>` resolves to a `DrawingFill` with `kind === :none` and a
`site`, because deliberately off is a setting.

`site` is the rung that answered. `chain` is every rung that was or could have
been consulted, highest precedence first, and is populated whether or not a
value was found — a setter uses it to decide where to write.
"""
struct Effective{T}
    value::Union{Nothing,T}
    site::Union{Nothing,FormatSite}
    chain::Vector{FormatSite}
end

"""
    SchemaKey

The `(namespace, complex-type)` pair identifying which `xsd:sequence` governs an
element's children. Not derivable from the element's tag: every series is `c:ser`
but its child order depends on the group containing it (`barSer`, `lineSer`, …),
and `c:spPr` is `a:CT_ShapeProperties` despite its chart prefix. Callers state it.
"""
const SchemaKey = Tuple{String,String}

# Value equality for the chart and DrawingML value types: every field except `raw`,
# which is a snapshot of the XML rather than part of the value. Fields compare with
# isequal, so `missing` (error cells in a cache) and NaN behave, and hash agrees --
# comparing with == would make -0.0 and 0.0 equal but hash differently.
# Chart and ChartEx are handles, compared by package and path in discovery.jl;
# ChartGroup is its key alone and SchemeColor normalizes its token, so both define
# their own.
for T in (ChartRef, ChartSeries, ChartAxis, ChartDataPoint, ChartDataLabel, ChartTrendline,
          ChartErrorBars, ChartMarker, ChartUpDownBars,
          ChartExDimension, ChartExData, ChartExBinning,
          DrawingColor, DrawingFill, DrawingLine, DrawingShapeProps, DrawingText,
          DrawingBodyProps, DrawingParagraph, DrawingParaProps, DrawingRun, DrawingRunProps)
    @eval begin
        Base.:(==)(a::$T, b::$T) =
            all(f -> f === :raw || isequal(getfield(a, f), getfield(b, f)), fieldnames($T))
        Base.hash(x::$T, h::UInt) =
            foldl((h, f) -> f === :raw ? h : hash(getfield(x, f), h), fieldnames($T);
                  init = hash($(QuoteNode(nameof(T))), h))
    end
end