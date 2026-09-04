# Appearance accessors for `c:` charts.
#
# This file sits above `drawingml.jl` in the layering: that file parses
# DrawingML — `spPr`, `txPr`, colours, fills, lines, text — for anything that
# carries it, knowing nothing about charts. This file knows the chart schema and
# finds the nodes: which element holds a series' fill, which axes a group plots
# against, where a per-point override lives. Nothing here parses DrawingML
# itself; it locates a node and hands it to `drawingml.jl`.
#
# Entry point. Every accessor starts from a `Chart`, which reaches its XML
# through `chart_root(c)` — `c.package` plus `c.path` into `xf.data`. Chart
# parts are parsed eagerly at open, so that lookup is a dict hit, and the node
# returned is the one `writexlsx` will serialize. Parsed structs keep their
# element in `raw`, so mutating through `raw` *is* the write; nothing needs to
# propagate. `Chart` objects are values parsed at a moment in time rather than
# live handles — `getCharts` re-parses on every call — but the nodes they hold
# are shared, so two `Chart`s from two calls edit the same tree.
#
# Indexing is 1-based throughout, over `Chart.series` and over data points as
# the user sees them. Excel's own identifiers — `c:idx`, `c:order`, `c:axId` —
# are kept on the structs but are not positions: they need not be contiguous,
# and a file where series or points were deleted will have gaps. `c:axId` is
# the exception that is genuinely useful as a key, since `c:crossAx` and a
# group's `c:axId` children reference it; `getChartAxis(c, axid)` looks up by it.
#
# Absent is not explicit. Every optional value is `Union{Nothing,T}`, and
# `nothing` means the element or attribute was not written — which in this
# schema means "inherit" or "Excel decides", never "off". An explicit
# `<a:noFill/>` means deliberately off, `<c:delete val="1"/>` means deliberately
# hidden. `has_fill`/`has_line` exist because conflating the two at a call site
# is easy and wrong. This matters for writing as much as reading: emitting a
# `val="0"` into a chart that never had one is a change to the file for no
# reason.
#
# Two ambiguities in the schema worth knowing before adding accessors here:
#
#   - `c:marker` is two different elements. Under `c:ser` or `c:dPt` it is
#     marker properties; under `c:lineChart` or `c:scatterChart` it is a boolean
#     show-markers flag. `_parse_marker` must only be called with the former.
#   - `c:dLbls` exists at three levels — series, group, and individual `c:dLbl`
#     — and they form an inheritance chain. A retyped label carries the same
#     formatting in both `c:tx/c:rich` and `c:txPr`, so a write must update both
#     or Excel shows the `rich` version and the edit appears to do nothing.
#
# The resolution of the inheritance cascade is stage 4 work. Nothing here
# returns a resolved value: an accessor reports what one node says, and
# `nothing` means that node is silent, not that a default applies.
#
# Cascade resolution deliberately stops at the chart XML. Where a property is
# written at no rung, `Effective.value` is `nothing`, which means Excel takes it
# from `xl/charts/style1.xml` and `colors1.xml`. Those parts are out of scope
# for stage 4 and are not read here.
#
# The reason is that Excel's precedence between them is underdocumented:
# style1.xml is keyed by a series-index-modulo scheme that varies by chart type,
# and colors1.xml layers on top. Pinning that down is research, and blocking the
# write path on it would be the wrong trade. `FormatSite.level` already accepts
# `:style` so adding the rung later is additive, not breaking.

const AXIS_TAGS = ("catAx", "valAx", "dateAx", "serAx")

const AXIS_ROLES = Dict(
    :category => "catAx",
    :value    => "valAx",
    :date     => "dateAx",
    :series   => "serAx",
)

# c:axId and c:crossAx carry unsigned 32-bit values, which Excel writes near the
# top of the range. Parse to Int (64-bit on every platform we support) rather
# than Int32 — see the 32-bit portability bug in table.jl.
function _axis_id(el::Union{Nothing,XML.Node}, tag::AbstractString)
    v = _attr(first_element_with_tag(el, tag), "val")
    return isnothing(v) ? nothing : parse(Int, v)
end

function parse_chart_axis(el::XML.Node)::ChartAxis
    pos = _attr(first_element_with_tag(el, "axPos"), "val")
    del = _attr(first_element_with_tag(el, "delete"), "val")
    return ChartAxis(
        Symbol(localname(el)),
        _axis_id(el, "axId"),
        isnothing(pos) ? nothing : Symbol(pos),
        _axis_id(el, "crossAx"),
        del in ("1", "true"),
        el,
    )
end

"""
    getChartAxes(c::Chart) -> Vector{ChartAxis}
    getChartAxes(c::Chart, role::Symbol) -> Vector{ChartAxis}

Axes of `c`, in `c:plotArea` document order. `role` is `:category`, `:value`,
`:date` or `:series` and filters by axis kind; a combo chart has two value axes,
so this returns a vector rather than one axis.

Axes with `c:delete val="1"` are included. Excel does not draw them, but they
remain in the XML and remain formattable — check `ax.deleted` to skip them.
"""
function getChartAxes(c::Chart)::Vector{ChartAxis}
    plotarea = first_element_with_tag(first_element_with_tag(chart_root(c), "chart"), "plotArea")
    isnothing(plotarea) && return ChartAxis[]
    return [parse_chart_axis(el) for el in XML.eachelement(plotarea)
            if localname(el) in AXIS_TAGS]
end

function getChartAxes(c::Chart, role::Symbol)::Vector{ChartAxis}
    haskey(AXIS_ROLES, role) || throw(XLSXError(
        "Unknown axis role `:$role`. Use one of " *
        join((":$k" for k in keys(AXIS_ROLES)), ", ") * "."))
    tag = AXIS_ROLES[role]
    return [ax for ax in getChartAxes(c) if String(ax.kind) == tag]
end

"""
    getChartAxis(c::Chart, axid::Integer) -> ChartAxis

The axis with `c:axId` equal to `axid`. This is the identifier chart groups and
`c:crossAx` use, so it is how you get from a series group, or from one axis, to
its partner.
"""
function getChartAxis(c::Chart, axid::Integer)::ChartAxis
    axes = getChartAxes(c)
    i = findfirst(ax -> ax.axid == axid, axes)
    isnothing(i) && throw(XLSXError(
        "Chart `$(c.name)` has no axis with axId $axid. Found: " *
        join((string(something(ax.axid, "?")) for ax in axes), ", ") * "."))
    return axes[i]
end

getAxisShapeProps(c::Chart, ax::ChartAxis) = parse_drawing_shape_props(_wb(c), ax.raw)

getAxisTextProps(c::Chart, ax::ChartAxis) = parse_drawing_text(_wb(c), ax.raw)

"""
    getAxisTitleRef(c::Chart, ax::ChartAxis) -> Union{Nothing,String}

The formula behind an axis title that references a cell
(`c:title/c:tx/c:strRef/c:f`). `nothing` for a literal or absent title.
"""
function getAxisTitleRef(c::Chart, ax::ChartAxis)
    tx = first_element_with_tag(first_element_with_tag(ax.raw, "title"), "tx")
    sr = first_element_with_tag(tx, "strRef")
    return isnothing(sr) ? nothing : child_text(sr, "f")
end

"""
    getAxisTitleText(c::Chart, ax::ChartAxis) -> Union{Nothing,DrawingText}

The axis title's text body (`c:title/c:tx/c:rich`). `nothing` means no title
element, or a title that references a cell rather than carrying literal text —
`c:tx/c:strRef` — which this does not resolve.
"""
function getAxisTitleText(c::Chart, ax::ChartAxis)
    tx = first_element_with_tag(first_element_with_tag(ax.raw, "title"), "tx")
    return parse_drawing_text(_wb(c), tx; tag="rich")
end

# An element whose presence turns a feature on, carrying an optional `c:spPr`.
#
# These are the three-state cases: absent means the feature is off, present with
# formatting means on and formatted, present without means on and drawn with
# inherited formatting. Returning `nothing` for the first and a
# `DrawingShapeProps` for both others keeps "is it on?" a single `isnothing`
# check at the call site, and the returned props still distinguish an absent
# fill from an explicit `<a:noFill/>` in the usual way.
function _optional_spPr(c::Chart, parent::Union{Nothing,XML.Node}, tag::AbstractString)
    el = first_element_with_tag(parent, tag)
    isnothing(el) && return nothing
    return something(parse_drawing_shape_props(_wb(c), el),
                     DrawingShapeProps(nothing, nothing, nothing, nothing, el))
end

"""
    getAxisGridlines(c::Chart, ax::ChartAxis; minor=false) -> Union{Nothing,DrawingShapeProps}

Line properties for the axis gridlines (`c:majorGridlines`, or
`c:minorGridlines` when `minor` is true).

`nothing` means no gridlines element, so none are drawn. An element present with
no `c:spPr` means gridlines are drawn with inherited formatting, and comes back
as a `DrawingShapeProps` with every field absent — so `isnothing` answers "are
there gridlines?" and the returned value answers "how are they formatted?".
"""
getAxisGridlines(c::Chart, ax::ChartAxis; minor::Bool=false) =
    _optional_spPr(c, ax.raw, minor ? "minorGridlines" : "majorGridlines")

function _require_kind(ax::ChartAxis, kinds::Tuple, what::AbstractString)
    ax.kind in kinds || throw(XLSXError(
        "$what is only defined on $(join(kinds, " or ")); this is a $(ax.kind)."))
end

"""
    getAxisPartner(c::Chart, ax::ChartAxis) -> Union{Nothing,ChartAxis}

The axis that `ax` crosses (`c:crossAx`). `nothing` if unwritten or dangling.
"""
function getAxisPartner(c::Chart, ax::ChartAxis)
    isnothing(ax.crossax) && return nothing
    i = findfirst(a -> a.axid == ax.crossax, getChartAxes(c))
    return isnothing(i) ? nothing : getChartAxes(c)[i]
end

_wb(c::AbstractChart) = get_workbook(c.package)

# Resolve a series index against a chart, with a message that says what went
# wrong. `c.series[i]` would throw a BoundsError, which is correct but tells an
# interactive user nothing about which chart or how many series it has.
function _series(c::Chart, i::Integer)
    n = length(c.series)
    n == 0 && throw(XLSXError("Chart `$(c.name)` has no series."))
    1 <= i <= n || throw(XLSXError(
        "Chart `$(c.name)` has $n series; asked for series $i."))
    return c.series[i]
end

"""
    getSeriesShapeProps(c::Chart, i::Integer) -> Union{Nothing,DrawingShapeProps}

Fill, line and effects for the graphic of series `i` (`c:ser/c:spPr`), where `i`
is a position in `c.series`, not the `c:idx` value. `nothing` means no `spPr`
was written, which in DrawingML means inherit from the chart style — not "no
formatting". Contrast an explicit `<a:noFill/>`, which `has_fill` reports as
deliberately off.
"""
getSeriesShapeProps(c::Chart, i::Integer) =
    parse_drawing_shape_props(_wb(c), _series(c, i).raw)

"""
    getSeriesLabelTextProps(c::Chart, i::Integer) -> Union{Nothing,DrawingText}

Text body for the data labels of series `i` (`c:ser/c:dLbls/c:txPr`). Per-point
overrides in `c:dLbls/c:dLbl` are not consulted; this is the series-level default.
"""
getSeriesLabelTextProps(c::Chart, i::Integer) =
    parse_drawing_text(_wb(c), first_element_with_tag(_series(c, i).raw, "dLbls"))

    # Shared by `getSeriesMarker` and `getDataPointMarker`. `parent` is a `c:ser` or
# `c:dPt` — anything that may carry a `c:marker` child.
#
# Note `c:marker` is ambiguous by tag: as a child of `c:lineChart` or
# `c:scatterChart` it is a boolean show-markers flag, not this element. Only
# call this with a series or data point.
function _parse_marker(wb::Workbook, parent::Union{Nothing,XML.Node})
    m = first_element_with_tag(parent, "marker")
    isnothing(m) && return nothing
    sym = _attr(first_element_with_tag(m, "symbol"), "val")
    sz  = _attr(first_element_with_tag(m, "size"), "val")
    return ChartMarker(isnothing(sym) ? nothing : Symbol(sym),
                       isnothing(sz)  ? nothing : parse(Int, sz),
                       parse_drawing_shape_props(wb, m),
                       m)
end

"""
    getSeriesMarker(c::Chart, i::Integer) -> Union{Nothing,ChartMarker}

Marker for series `i` of a line, scatter or radar chart (`c:ser/c:marker`).
`nothing` means no `c:marker` element; a marker whose `symbol` is `:none` is a
marker explicitly turned off, which is a different thing.
"""
getSeriesMarker(c::Chart, i::Integer) = _parse_marker(_wb(c), _series(c, i).raw)

"""
    getChartGroups(c::Chart) -> Vector{ChartGroup}

The chart-type groups in `c:plotArea`, in document order. A combo chart has
several; a plain chart has one. Use [`getGroupAxes`](@ref) to find which axes a
group's series are plotted against.
"""
function getChartGroups(c::Chart)::Vector{ChartGroup}
    plotarea = first_element_with_tag(
        first_element_with_tag(chart_root(c), "chart"), "plotArea")
    isnothing(plotarea) && return ChartGroup[]

    groups = ChartGroup[]
    for el in XML.eachelement(plotarea)
        localname(el) in CHART_GROUP_TAGS || continue
        ids = Int[]
        for a in elements_with_tag(el, "axId")
            v = _attr(a, "val")
            isnothing(v) || push!(ids, parse(Int, v))
        end
        push!(groups, ChartGroup(Symbol(localname(el)), ids, el))
    end
    return groups
end

"""
    getGroupAxes(c::Chart, g::ChartGroup) -> Vector{ChartAxis}

The axes named by `g`'s `c:axId` children, in that order. Empty for pie and
doughnut groups, which have no axes. An id that names no axis in the plot area
is skipped rather than raising — malformed files exist.
"""
function getGroupAxes(c::Chart, g::ChartGroup)::Vector{ChartAxis}
    axes = getChartAxes(c)
    out = ChartAxis[]
    for id in g.axids
        i = findfirst(ax -> ax.axid == id, axes)
        isnothing(i) || push!(out, axes[i])
    end
    return out
end

"""
    getGroupLabelTextProps(c::Chart, g::ChartGroup) -> Union{Nothing,DrawingText}

Group-level data label text (`c:barChart/c:dLbls/c:txPr`). This is the
inheritance parent for the series-level `c:dLbls`; Excel commonly writes a group
`c:dLbls` carrying only the `show*` flags and no `txPr`, in which case this is
`nothing`.
"""
getGroupLabelTextProps(c::Chart, g::ChartGroup) =
    parse_drawing_text(_wb(c), first_element_with_tag(g.raw, "dLbls"))

"""
    getSeriesGroup(c::Chart, i::Integer) -> ChartGroup

The chart-type group containing series `i`.
"""
function getSeriesGroup(c::Chart, i::Integer)::ChartGroup
    ser = _series(c, i).raw
    for g in getChartGroups(c)
        any(s -> s === ser, elements_with_tag(g.raw, "ser")) && return g
    end
    throw(XLSXError("Series $i of chart `$(c.name)` is not in any chart group."))
end

"""
    getSeriesAxes(c::Chart, i::Integer) -> Vector{ChartAxis}

The axes series `i` is plotted against. This is how you tell a primary-axis
series from a secondary-axis one: they sit in different groups naming different
axis ids.
"""
getSeriesAxes(c::Chart, i::Integer) = getGroupAxes(c, getSeriesGroup(c, i))

"""
    getChartLegend(c::Chart) -> Union{Nothing,XML.Node}

The `c:legend` element, or `nothing` if the chart has no legend. Returned as a
node because the legend carries no identity of its own — pass it to
[`getLegendShapeProps`](@ref) and friends, or read it directly.
"""
getChartLegend(c::Chart) =
    first_element_with_tag(first_element_with_tag(chart_root(c), "chart"), "legend")

"""
    getLegendPos(c::Chart) -> Union{Nothing,Symbol}

`:b`, `:t`, `:l`, `:r` or `:tr` from `c:legendPos`. `nothing` means no legend,
or a legend with no explicit position.
"""
function getLegendPos(c::Chart)
    v = _attr(first_element_with_tag(getChartLegend(c), "legendPos"), "val")
    return isnothing(v) ? nothing : Symbol(v)
end

getLegendOverlay(c::Chart) = _bool_val(getChartLegend(c), "overlay")

getLegendShapeProps(c::Chart) = parse_drawing_shape_props(_wb(c), getChartLegend(c))
getLegendTextProps(c::Chart)  = parse_drawing_text(_wb(c), getChartLegend(c))

_chartnode(c::Chart) = first_element_with_tag(chart_root(c), "chart")

getChartTitleNode(c::Chart) = first_element_with_tag(_chartnode(c), "title")

"""
    getChartTitleText(c::Chart) -> Union{Nothing,DrawingText}

Literal title text (`c:title/c:tx/c:rich`). `nothing` if there is no title
element, or the title references a cell rather than carrying literal text — see
[`getChartTitleRef`](@ref) — or `c:autoTitleDeleted` is 0 and Excel generates the
title from the series name, in which case no text exists in the file at all.
"""
getChartTitleText(c::Chart) =
    parse_drawing_text(_wb(c),
        first_element_with_tag(getChartTitleNode(c), "tx"); tag="rich")

"""
    getChartTitleRef(c::Chart) -> Union{Nothing,String}

The formula behind a title that references a cell (`c:title/c:tx/c:strRef/c:f`).
`nothing` for a literal or absent title.
"""
function getChartTitleRef(c::Chart)
    tx = first_element_with_tag(getChartTitleNode(c), "tx")
    sr = first_element_with_tag(tx, "strRef")
    return isnothing(sr) ? nothing : child_text(sr, "f")
end

getChartTitleShapeProps(c::Chart) = parse_drawing_shape_props(_wb(c), getChartTitleNode(c))
getChartTitleTextProps(c::Chart)  = parse_drawing_text(_wb(c), getChartTitleNode(c))

getAutoTitleDeleted(c::Chart) = _bool_val(_chartnode(c), "autoTitleDeleted")

getPlotAreaShapeProps(c::Chart) =
    parse_drawing_shape_props(_wb(c), first_element_with_tag(_chartnode(c), "plotArea"))

getChartSpaceShapeProps(c::Chart) = parse_drawing_shape_props(_wb(c), chart_root(c))
getChartSpaceTextProps(c::Chart)  = parse_drawing_text(_wb(c), chart_root(c))

# --- axis scalars ---------------------------------------------------------
#
# All of these are elements carrying a `val` attribute, not attributes on the
# axis — the "elements masquerading as attributes" pattern. `nothing` means the
# element was not written, which for an axis means Excel applies its own default
# rather than that the feature is off.

# Read a `val` attribute from a child element of `ax`, as a Symbol.
function _axis_sym(ax::ChartAxis, tag::AbstractString)
    v = _attr(first_element_with_tag(ax.raw, tag), "val")
    return isnothing(v) ? nothing : Symbol(v)
end

# Read a `val` attribute from a child element of `ax`, as a Float64.
# Numeric axis values are xsd:double throughout; dates on a dateAx arrive as
# serial numbers, matching how the rest of the package handles date cells.
function _axis_num(ax::ChartAxis, tag::AbstractString)
    v = _attr(first_element_with_tag(ax.raw, tag), "val")
    return isnothing(v) ? nothing : parse(Float64, v)
end

# As `_axis_num`, but reading from a child of `c:scaling` rather than the axis.
function _scaling_num(sc::XML.Node, tag::AbstractString)
    v = _attr(first_element_with_tag(sc, tag), "val")
    return isnothing(v) ? nothing : parse(Float64, v)
end

"""
    getAxisNumberFormatCode(c::Chart, ax::ChartAxis) -> Union{Nothing,String}

The format code from `c:numFmt/@formatCode`. Ignored by Excel when
[`getAxisNumberFormatLinked`](@ref) is `true`.
"""
getAxisNumberFormatCode(c::Chart, ax::ChartAxis) =
    _attr(first_element_with_tag(ax.raw, "numFmt"), "formatCode")

"""
    getAxisNumberFormatLinked(c::Chart, ax::ChartAxis) -> Union{Nothing,Bool}

`c:numFmt/@sourceLinked`. `true` means Excel takes the number format from the
source cells and ignores the axis's own format code. `nothing` means no
`c:numFmt` element was written.
"""
function getAxisNumberFormatLinked(c::Chart, ax::ChartAxis)
    el = first_element_with_tag(ax.raw, "numFmt")
    isnothing(el) && return nothing
    return _attr(el, "sourceLinked") in ("1", "true")
end

"""
    getAxisMajorTickMark(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:majorTickMark` — `:cross`, `:in`, `:out` or `:none`.
"""
getAxisMajorTickMark(c::Chart, ax::ChartAxis) = _axis_sym(ax, "majorTickMark")

"""
    getAxisMinorTickMark(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:minorTickMark` — `:cross`, `:in`, `:out` or `:none`.
"""
getAxisMinorTickMark(c::Chart, ax::ChartAxis) = _axis_sym(ax, "minorTickMark")

"""
    getAxisTickLabelPos(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:tickLblPos` — `:high`, `:low`, `:nextTo` or `:none`.
"""
getAxisTickLabelPos(c::Chart, ax::ChartAxis) = _axis_sym(ax, "tickLblPos")

"""
    getAxisOrientation(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:scaling/c:orientation` — `:minMax` for a normal axis, `:maxMin` for one
plotted in reverse order.

Throws if the axis has no `c:scaling`: the schema requires it, so its absence
means a malformed chart part rather than an inherited value.
"""
getAxisOrientation(c::Chart, ax::ChartAxis) =
    (v = _attr(first_element_with_tag(_axis_scaling(ax), "orientation"), "val");
     isnothing(v) ? nothing : Symbol(v))

"""
    getAxisMin(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Fixed lower bound from `c:scaling/c:min`. `nothing` means Excel scales the axis
automatically, which is the usual case; a value means the user fixed the bound.
"""
getAxisMin(c::Chart, ax::ChartAxis) = _scaling_num(_axis_scaling(ax), "min")

"""
    getAxisMax(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Fixed upper bound from `c:scaling/c:max`. `nothing` means automatic.
"""
getAxisMax(c::Chart, ax::ChartAxis) = _scaling_num(_axis_scaling(ax), "max")

"""
    getAxisLogBase(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

`c:scaling/c:logBase`. `nothing` means a linear axis.
"""
getAxisLogBase(c::Chart, ax::ChartAxis) = _scaling_num(_axis_scaling(ax), "logBase")

# c:scaling is required by CT_Scaling's parent in the schema, so a missing one
# is a malformed part, not an absent-means-inherit case.
function _axis_scaling(ax::ChartAxis)
    sc = first_element_with_tag(ax.raw, "scaling")
    isnothing(sc) && throw(XLSXError(
        "Axis $(something(ax.axid, "?")) ($(ax.kind)) has no `c:scaling` element, " *
        "which the schema requires. The chart part is malformed."))
    return sc
end

"""
    getAxisCrosses(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

Where the partner axis crosses this one, as a rule: `:autoZero`, `:min` or
`:max`. Mutually exclusive with [`getAxisCrossesAt`](@ref) — a chart uses one or
the other, so at most one of the two is non-`nothing`.
"""
getAxisCrosses(c::Chart, ax::ChartAxis) = _axis_sym(ax, "crosses")

"""
    getAxisCrossesAt(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Where the partner axis crosses this one, as a value (`c:crossesAt`). Mutually
exclusive with [`getAxisCrosses`](@ref).
"""
getAxisCrossesAt(c::Chart, ax::ChartAxis) = _axis_num(ax, "crossesAt")

"""
    getAxisMajorUnit(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Interval between major ticks and gridlines (`c:majorUnit`). `nothing` means
Excel chooses automatically. Value and date axes only.
"""
function getAxisMajorUnit(c::Chart, ax::ChartAxis)
    _require_kind(ax, (:valAx, :dateAx), "majorUnit")
    return _axis_num(ax, "majorUnit")
end

"""
    getAxisMinorUnit(c::Chart, ax::ChartAxis) -> Union{Nothing,Float64}

Interval between minor ticks (`c:minorUnit`). Value and date axes only.
"""
function getAxisMinorUnit(c::Chart, ax::ChartAxis)
    _require_kind(ax, (:valAx, :dateAx), "minorUnit")
    return _axis_num(ax, "minorUnit")
end

"""
    getAxisLabelAlign(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:lblAlgn` — `:ctr`, `:l` or `:r`. Category and date axes only.
"""
function getAxisLabelAlign(c::Chart, ax::ChartAxis)
    _require_kind(ax, (:catAx, :dateAx), "lblAlgn")
    return _axis_sym(ax, "lblAlgn")
end

"""
    getAxisLabelOffset(c::Chart, ax::ChartAxis) -> Union{Nothing,Int}

`c:lblOffset` — distance of the labels from the axis, as a percentage between 0
and 1000. Category and date axes only.
"""
function getAxisLabelOffset(c::Chart, ax::ChartAxis)
    _require_kind(ax, (:catAx, :dateAx), "lblOffset")
    v = _attr(first_element_with_tag(ax.raw, "lblOffset"), "val")
    return isnothing(v) ? nothing : parse(Int, v)
end

"""
    getAxisMultiLevelLabels(c::Chart, ax::ChartAxis) -> Union{Nothing,Bool}

Whether multi-level category labels are shown. This inverts the XML: the
element is `c:noMultiLvlLbl`, so `val="1"` means labels are *not* multi-level
and this returns `false`. Category and date axes only.
"""
function getAxisMultiLevelLabels(c::Chart, ax::ChartAxis)
    _require_kind(ax, (:catAx, :dateAx), "noMultiLvlLbl")
    v = _bool_val(ax.raw, "noMultiLvlLbl")
    return isnothing(v) ? nothing : !v
end


"""
    getAxisCrossBetween(c::Chart, ax::ChartAxis) -> Union{Nothing,Symbol}

`c:crossBetween` — `:between` if the partner axis crosses between categories,
`:midCat` if it crosses at their midpoints. Value axes only.
"""
function getAxisCrossBetween(c::Chart, ax::ChartAxis)
    _require_kind(ax, (:valAx,), "crossBetween")
    return _axis_sym(ax, "crossBetween")
end

"""
    getSeriesDataPoints(c::Chart, i::Integer) -> Vector{ChartDataPoint}

Per-point overrides on series `i`, in document order. Usually empty, or short —
only points the user formatted individually appear.
"""
function getSeriesDataPoints(c::Chart, i::Integer)::Vector{ChartDataPoint}
    out = ChartDataPoint[]
    for el in elements_with_tag(_series(c, i).raw, "dPt")
        n = _attr(first_element_with_tag(el, "idx"), "val")
        isnothing(n) && continue          # a dPt with no idx applies to nothing
        push!(out, ChartDataPoint(
            parse(Int, n),
            _bool_val(el, "invertIfNegative"),
            _bool_val(el, "bubble3D"),
            el))
    end
    return out
end

# Read a boolean `val` attribute from a child element. Returns `nothing` when
# the element is absent, which throughout the chart schema means "inherit" or
# "Excel's default" rather than false.
_bool_val(el::Union{Nothing,XML.Node}, tag::AbstractString) =
    (v = _attr(first_element_with_tag(el, tag), "val");
     isnothing(v) ? nothing : v in ("1", "true"))

_int_val(el::Union{Nothing,XML.Node}, tag::AbstractString) =
    (v = _attr(first_element_with_tag(el, tag), "val");
     isnothing(v) ? nothing : parse(Int, v))

_num_val(el::Union{Nothing,XML.Node}, tag::AbstractString) =
    (v = _attr(first_element_with_tag(el, tag), "val");
     isnothing(v) ? nothing : parse(Float64, v))

_sym_val(el::Union{Nothing,XML.Node}, tag::AbstractString) =
    (v = _attr(first_element_with_tag(el, tag), "val");
     isnothing(v) ? nothing : Symbol(v))

"""
    getSeriesDataPoint(c::Chart, i::Integer, point::Integer) -> Union{Nothing,ChartDataPoint}

The override for data point `point` of series `i`, counting from 1, or
`nothing` if that point has no individual formatting.
"""
function getSeriesDataPoint(c::Chart, i::Integer, point::Integer)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    dps = getSeriesDataPoints(c, i)
    j = findfirst(d -> d.idx == point - 1, dps)
    return isnothing(j) ? nothing : dps[j]
end

getDataPointShapeProps(c::Chart, d::ChartDataPoint) =
    parse_drawing_shape_props(_wb(c), d.raw)

"""
    getDataPointMarker(c::Chart, d::ChartDataPoint) -> Union{Nothing,ChartMarker}

Marker override for a single data point (`c:dPt/c:marker`). `nothing` means the
point does not override the series marker — the series-level
[`getSeriesMarker`](@ref) applies.
"""
getDataPointMarker(c::Chart, d::ChartDataPoint) = _parse_marker(_wb(c), d.raw)

# The `c:dLbls` container on a series, group or data point.
_dlbls(parent::Union{Nothing,XML.Node}) = first_element_with_tag(parent, "dLbls")

"""
    getSeriesDataLabels(c::Chart, i::Integer) -> Vector{ChartDataLabel}

Individual label overrides on series `i`, in document order. Usually empty.
"""
function getSeriesDataLabels(c::Chart, i::Integer)::Vector{ChartDataLabel}
    out = ChartDataLabel[]
    dl = _dlbls(_series(c, i).raw)
    isnothing(dl) && return out
    for el in elements_with_tag(dl, "dLbl")
        n = _attr(first_element_with_tag(el, "idx"), "val")
        isnothing(n) && continue
        push!(out, ChartDataLabel(parse(Int, n),
                            _bool_val(el, "delete"),
                            el))
    end
    return out
end

"""
    getSeriesDataLabel(c::Chart, i::Integer, point::Integer) -> Union{Nothing,ChartDataLabel}

The label override for data point `point` of series `i`, counting from 1, or
`nothing` if that label is not individually formatted.
"""
function getSeriesDataLabel(c::Chart, i::Integer, point::Integer)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    dls = getSeriesDataLabels(c, i)
    j = findfirst(d -> d.idx == point - 1, dls)
    return isnothing(j) ? nothing : dls[j]
end

getDataLabelTextProps(c::Chart, d::ChartDataLabel) =
    parse_drawing_text(_wb(c), d.raw)

getDataLabelShapeProps(c::Chart, d::ChartDataLabel) =
    parse_drawing_shape_props(_wb(c), d.raw)

"""
    getDataLabelText(c::Chart, d::ChartDataLabel) -> Union{Nothing,DrawingText}

Literal replacement text for one label (`c:dLbl/c:tx/c:rich`). `nothing` means
the label shows its value rather than typed-over text. Contrast
[`getDataLabelTextProps`](@ref), which is the label's `c:txPr` formatting.
"""
getDataLabelText(c::Chart, d::ChartDataLabel) =
    parse_drawing_text(_wb(c), first_element_with_tag(d.raw, "tx"); tag="rich")

"""
    getDataLabelPosition(c::Chart, d::ChartDataLabel) -> Union{Nothing,Symbol}

`c:dLblPos` — `:ctr`, `:inEnd`, `:inBase`, `:outEnd`, `:l`, `:r`, `:t`, `:b`,
`:bestFit`. Which values are legal depends on the chart type.
"""
function getDataLabelPosition(c::Chart, d::ChartDataLabel)
    v = _attr(first_element_with_tag(d.raw, "dLblPos"), "val")
    return isnothing(v) ? nothing : Symbol(v)
end

"""
    getDataLabelOffset(c::Chart, d::ChartDataLabel) -> Union{Nothing,NamedTuple}

Manual position offset (`c:dLbl/c:layout/c:manualLayout`) as `(x, y)`,
fractions of the plot area, signed. `nothing` means the label sits where Excel
puts it. Distinct from [`getDataLabelPosition`](@ref), which is the discrete
`c:dLblPos` enum — Excel writes a layout offset when you drag a label and a
position when you pick one from the menu, so a label usually has one or neither.
"""
function getDataLabelOffset(c::Chart, d::ChartDataLabel)
    ml = first_element_with_tag(first_element_with_tag(d.raw, "layout"), "manualLayout")
    isnothing(ml) && return nothing
    f(tag) = (v = _attr(first_element_with_tag(ml, tag), "val");
              isnothing(v) ? nothing : parse(Float64, v))
    return (x = f("x"), y = f("y"))
end



    # --- trendlines -----------------------------------------------------------

"""
    getSeriesTrendlines(c::Chart, i::Integer) -> Vector{ChartTrendline}

Trendlines on series `i`, in document order. A series can carry several — a
linear fit and a moving average, say — so this returns a vector. Usually empty.
"""
function getSeriesTrendlines(c::Chart, i::Integer)::Vector{ChartTrendline}
    out = ChartTrendline[]
    for el in elements_with_tag(_series(c, i).raw, "trendline")
        t = _attr(first_element_with_tag(el, "trendlineType"), "val")
        push!(out, ChartTrendline(
            isnothing(t) ? nothing : Symbol(t),
            child_text(el, "name"),
            _int_val(el, "order"),
            _int_val(el, "period"),
            _num_val(el, "forward"),
            _num_val(el, "backward"),
            _num_val(el, "intercept"),
            _bool_val(el, "dispRSqr"),
            _bool_val(el, "dispEq"),
            el))
    end
    return out
end

"""
    getTrendlineShapeProps(c::Chart, t::ChartTrendline) -> Union{Nothing,DrawingShapeProps}

Line properties of the trendline itself (`c:trendline/c:spPr`). Trendlines are
lines, so the fill is normally absent; the dash pattern lives in `line.dash`.
"""
getTrendlineShapeProps(c::Chart, t::ChartTrendline) =
    parse_drawing_shape_props(_wb(c), t.raw)

"""
    getTrendlineLabelText(c::Chart, t::ChartTrendline) -> Union{Nothing,DrawingText}

Literal text of the trendline's on-chart label (`c:trendlineLbl/c:tx/c:rich`).
`nothing` means the label shows Excel's generated equation or R² rather than
typed-over text — or that there is no label at all. Contrast
[`getTrendlineLabelTextProps`](@ref), which is the label's `c:txPr` formatting.
"""
getTrendlineLabelText(c::Chart, t::ChartTrendline) =
    parse_drawing_text(_wb(c),
        first_element_with_tag(_trendline_lbl(t), "tx"); tag="rich")

"""
    getTrendlineLabelTextProps(c::Chart, t::ChartTrendline) -> Union{Nothing,DrawingText}

Formatting of the trendline label (`c:trendlineLbl/c:txPr`).
"""
getTrendlineLabelTextProps(c::Chart, t::ChartTrendline) =
    parse_drawing_text(_wb(c), _trendline_lbl(t))

"""
    getTrendlineLabelShapeProps(c::Chart, t::ChartTrendline) -> Union{Nothing,DrawingShapeProps}

Fill and border of the trendline label's box (`c:trendlineLbl/c:spPr`).
"""
getTrendlineLabelShapeProps(c::Chart, t::ChartTrendline) =
    parse_drawing_shape_props(_wb(c), _trendline_lbl(t))

_trendline_lbl(t::ChartTrendline) = first_element_with_tag(t.raw, "trendlineLbl")

# --- error bars -----------------------------------------------------------

"""
    getSeriesErrorBars(c::Chart, i::Integer) -> Vector{ChartErrorBars}

Error bars on series `i`, in document order. A series carries at most two, one
for each of the `x` and `y` directions.
"""
function getSeriesErrorBars(c::Chart, i::Integer)::Vector{ChartErrorBars}
    out = ChartErrorBars[]
    for el in elements_with_tag(_series(c, i).raw, "errBars")
        push!(out, ChartErrorBars(
            _sym_val(el, "errDir"),
            _sym_val(el, "errBarType"),
            _sym_val(el, "errValType"),
            _num_val(el, "val"),
            _bool_val(el, "noEndCap"),
            el))
    end
    return out
end

"""
    getErrorBarsShapeProps(c::Chart, e::ChartErrorBars) -> Union{Nothing,DrawingShapeProps}

Line properties of the error bars (`c:errBars/c:spPr`).
"""
getErrorBarsShapeProps(c::Chart, e::ChartErrorBars) =
    parse_drawing_shape_props(_wb(c), e.raw)

"""
    getErrorBarsCustomRefs(c::Chart, e::ChartErrorBars) -> NamedTuple

The sheet references behind custom error bars, as `(plus, minus)`, each a
[`ChartRef`](@ref) or `nothing`.

Only `errValType="cust"` bars carry these; every other value type takes its
magnitude from [`ChartErrorBars`](@ref)`.value` instead, and this returns
`(nothing, nothing)`. Excel writes both `c:plus` and `c:minus` even for
one-sided bars, so a `nothing` here means the element was genuinely absent.

Cached values are read, matching `read_cached_values=true` on
[`getCharts`](@ref) — these ranges are short, so there is nothing to save by
skipping them.
"""
function getErrorBarsCustomRefs(c::Chart, e::ChartErrorBars)
    f(tag) = parse_chart_ref(first_element_with_tag(e.raw, tag);
                             read_cached_values=true)
    return (plus = f("plus"), minus = f("minus"))
end

# --- line and stock chart extras ------------------------------------------
#
# These hang off the group rather than a series: drop lines run from a point to
# the axis, hi-low lines and up-down bars span between series, so none of them
# belongs to one series. Each carries only `c:spPr`.
#
# Presence is the state. Unlike most of this file, `nothing` here means the
# feature is off rather than inherited — Excel writes the element when you turn
# the feature on and removes it when you turn it off. An element present with no
# `c:spPr` means on, drawn with inherited formatting.

"""
    getGroupDropLines(c::Chart, g::ChartGroup) -> Union{Nothing,DrawingShapeProps}

Line properties of the group's drop lines (`c:dropLines`), which run from each
data point down to the category axis. `nothing` means the group has no drop
lines. Line, area and stock charts.
"""
getGroupDropLines(c::Chart, g::ChartGroup) = _optional_spPr(c, g.raw, "dropLines")

"""
    getGroupHiLowLines(c::Chart, g::ChartGroup) -> Union{Nothing,DrawingShapeProps}

Line properties of the group's high-low lines (`c:hiLowLines`), which span
between the highest and lowest series at each category. Line and stock charts.
"""
getGroupHiLowLines(c::Chart, g::ChartGroup) = _optional_spPr(c, g.raw, "hiLowLines")

"""
    getGroupSeriesLines(c::Chart, g::ChartGroup) -> Union{Nothing,DrawingShapeProps}

Line properties of the group's series lines (`c:serLines`), which connect
segments across categories on a stacked bar or an of-pie chart.
"""
getGroupSeriesLines(c::Chart, g::ChartGroup) = _optional_spPr(c, g.raw, "serLines")

function getGroupUpDownBars(c::Chart, g::ChartGroup)
    el = first_element_with_tag(g.raw, "upDownBars")
    isnothing(el) && return nothing
    return ChartUpDownBars(_int_val(el, "gapWidth"), el)
end

"""
    getUpBarShapeProps(c::Chart, b::ChartUpDownBars) -> Union{Nothing,DrawingShapeProps}

Fill and border of the up bars (`c:upBars/c:spPr`). `nothing` means `c:upBars`
was written with no formatting, or not written at all — Excel writes an empty
`<c:upBars/>` for default white bars, so check `b.raw` if the distinction
matters.
"""
getUpBarShapeProps(c::Chart, b::ChartUpDownBars) =
    parse_drawing_shape_props(_wb(c), first_element_with_tag(b.raw, "upBars"))

getDownBarShapeProps(c::Chart, b::ChartUpDownBars) =
    parse_drawing_shape_props(_wb(c), first_element_with_tag(b.raw, "downBars"))

    #-- Cascade resolution -------------------------------------------------------

const _RUN_PROP_TYPES = Dict{Symbol,Type}(
    f => Base.nonnothingtype(t)
    for (f, t) in zip(fieldnames(DrawingRunProps), fieldtypes(DrawingRunProps))
    if f !== :raw
)
const _RUN_PROP_FIELDS = Tuple(f for f in fieldnames(DrawingRunProps) if haskey(_RUN_PROP_TYPES, f))

"""
    _walk_fill(wb, chain) -> Effective{DrawingFill}

Return the first fill written on any rung of `chain`, with the rung that
carried it. A rung whose `spPr` exists but holds no fill element is skipped —
a series with only a line inherits its fill from the next rung down.
"""
function _walk_fill(wb::Workbook, chain::Vector{FormatSite})::Effective{DrawingFill}
    for s in chain
        isnothing(s.props) && continue
        f = parse_drawing_fill(wb, s.props)
        isnothing(f) || return Effective{DrawingFill}(f, s, chain)
    end
    Effective{DrawingFill}(nothing, nothing, chain)
end

function _walk_line(wb::Workbook, chain::Vector{FormatSite})::Effective{DrawingLine}
    for s in chain
        isnothing(s.props) && continue
        l = parse_drawing_line(wb, s.props)
        isnothing(l) || return Effective{DrawingLine}(l, s, chain)
    end
    Effective{DrawingLine}(nothing, nothing, chain)
end

"""
    _shape_chain(c, i; point=nothing) -> Vector{FormatSite}

The `spPr` cascade for the graphic of series `i`: the data point, then the
series. Chart groups carry no `spPr`, and `plotArea`/`chartSpace` `spPr`
describe their own backgrounds rather than a series, so the chain stops there
and falls through to the style part, which this layer does not read.
"""
function _shape_chain(c::Chart, i::Integer; point::Union{Nothing,Integer}=nothing)
    chain = FormatSite[]
    if !isnothing(point)
        dp = getSeriesDataPoint(c, i, point)
        isnothing(dp) || push!(chain, _site_path(:point, :shape, dp.raw, "spPr"))
    end
    push!(chain, _site_path(:series, :shape, _series(c, i).raw, "spPr"))
    return chain
end

"""
    getSeriesFill(c::Chart, i::Integer; point=nothing) -> Effective{DrawingFill}

Resolve the fill of the graphic of series `i`, or of one data point of it.

This is always `c:spPr`. On a line or scatter series what Excel's format pane
calls the point's fill lives on its marker instead — use
[`getMarkerFill`](@ref) for that.
"""
getSeriesFill(c::Chart, i::Integer; point::Union{Nothing,Integer}=nothing) =
    _walk_fill(_wb(c), _shape_chain(c, i; point))

"""
    getSeriesLine(c::Chart, i::Integer; point=nothing) -> Effective{DrawingLine}

Resolve the outline of the graphic of series `i`, or of one data point of it.
"""
getSeriesLine(c::Chart, i::Integer; point::Union{Nothing,Integer}=nothing) =
    _walk_line(_wb(c), _shape_chain(c, i; point))

"""
    getMarkerFill(c::Chart, i::Integer, point::Integer) -> Effective{DrawingFill}

Resolve the fill of a data point's marker: `c:dPt/c:marker/c:spPr`, then
`c:ser/c:marker/c:spPr`.

Both lookups are anchored on `c:ser` and `c:dPt`, where `c:marker` is marker
properties. Under a group element the same tag is a boolean show-markers flag,
so this must not be generalised to walk up from a group.
"""
function getMarkerFill(c::Chart, i::Integer, point::Integer)::Effective{DrawingFill}
    chain = FormatSite[]
    dp = getSeriesDataPoint(c, i, point)
    isnothing(dp) || push!(chain, _site_path(:point, :marker, dp.raw, "marker", "spPr"))
    push!(chain, _site_path(:series, :marker, _series(c, i).raw, "marker", "spPr"))
    _walk_fill(_wb(c), chain)
end

"""
    getLabelTextProp(c::Chart, i::Integer, field::Symbol) -> Effective

Resolve one field of the data-label text formatting for series `i`, walking
`c:txPr` from the series up through the group, plot area and chart space.

`field` is a field of `DrawingRunProps` other than `raw`. Text inherits field by
field — a size from one rung and a weight from another both apply — which is
why this resolves a named field rather than a whole `DrawingRunProps`.

Two fields are composite: `:fill` and `:line` resolve to a whole `DrawingFill`
or `DrawingLine`. They inherit as units — a run's fill is written at some rung
or it is not — so there is no cascade below them, and resolving through to
`.fgcolor` or `.width` is a field access on the result, not a further walk.

A rung whose `txPr` exists but sets nothing (an empty `a:defRPr`, which Excel
writes at the chart space) does not answer: the walk continues, and a field set
at no rung resolves to `nothing`, meaning Excel takes it from the chart style.
"""
function getLabelTextProp(c::Chart, i::Integer, field::Symbol)
    field in _RUN_PROP_FIELDS || throw(XLSXError(
        "`$field` is not a resolvable text property. Valid fields: " *
        join(_RUN_PROP_FIELDS, ", ") * "."))

    wb = _wb(c)
    chain = _text_chain(c, i)
    for s in chain
        isnothing(s.props) && continue
        txt = parse_drawing_text(wb, s.props)
        isnothing(txt) && continue
        rp = default_run_props(txt)
        isnothing(rp) && continue          # mixed text — no single answer at this rung
        v = getfield(rp, field)
        isnothing(v) || return Effective{_RUN_PROP_TYPES[field]}(v, s, chain)
    end
    return Effective{_RUN_PROP_TYPES[field]}(nothing, nothing, chain)
end

"""
    _site_path(level, kind, container, path...) -> FormatSite

Build a rung whose properties live at `path` below `container`. `container` is
the nearest element that always exists — the `c:ser`, not its `c:dLbls`, which
Excel omits when a series has no labels. A setter creates whatever the path
lacks; `props === nothing` means it lacks something.
"""
function _site_path(level::Symbol, kind::Symbol, container::XML.Node, path::AbstractString...)
    n = container
    for tag in path
        n = first_element_with_tag(n, tag)
        isnothing(n) && return FormatSite(level, kind, container, nothing)
    end
    return FormatSite(level, kind, container, n)
end

function _text_chain(c::Chart, i::Integer)
    root = chart_root(c)
    chain = FormatSite[
        _site_path(:series, :text, _series(c, i).raw,      "dLbls", "txPr"),
        _site_path(:group,  :text, getSeriesGroup(c, i).raw, "dLbls", "txPr"),
    ]
    chrt = first_element_with_tag(root, "chart")
    pa = isnothing(chrt) ? nothing : first_element_with_tag(chrt, "plotArea")
    isnothing(pa) || push!(chain, _site_path(:plotarea, :text, pa, "txPr"))
    push!(chain, _site_path(:chartspace, :text, root, "txPr"))
    return chain
end