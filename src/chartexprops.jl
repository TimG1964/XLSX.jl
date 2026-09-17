# Accessors for ChartEx: charts in the Microsoft cx: namespace
# (http://schemas.microsoft.com/office/drawing/2014/chartex).
#
# A ChartEx holds identity only; everything here reads the current part, so
# results never go stale. Series are indexed 1-based over cx:series in document
# order. A Pareto chart is written as column series each owned by a paretoLine
# series, and both count.
#
# Discovery readers (_cx_root, _cx_chart, _cx_series_nodes, _cx_layouts,
# _cx_has_binning, _cx_refs, _cx_title) live in charts.jl.

# ---- small readers ---------------------------------------------------------

function _cx_series_node(c::ChartEx, i::Integer)
    s = _cx_series_nodes(c)
    1 <= i <= length(s) ||
        throw(XLSXError("Series $i out of range: chart `$(c.name)` has $(length(s)) series."))
    return s[i]
end

_cx_layoutpr(c::ChartEx, i::Integer) = first_element_with_tag(_cx_series_node(c, i), "layoutPr")
_cx_datalabels(c::ChartEx, i::Integer) = first_element_with_tag(_cx_series_node(c, i), "dataLabels")

# Attribute value as written, or `nothing` when the node or attribute is absent.
_cx_attr(node, name) = isnothing(node) ? nothing : (s = get_attr(node, name); isempty(s) ? nothing : s)

_cx_symbol(node, name) = (s = _cx_attr(node, name); isnothing(s) ? nothing : Symbol(s))

function _cx_bool(node, name)::Union{Nothing,Bool}
    s = _cx_attr(node, name)
    isnothing(s) && return nothing
    s in ("1", "true")  && return true
    s in ("0", "false") && return false
    throw(XLSXError("Attribute `$name` should be a boolean. Found \"$s\"."))
end

function _cx_num_or_auto(node, name, T::Type)
    s = _cx_attr(node, name)
    isnothing(s) && return nothing
    s == "auto" && return :auto
    return parse(T, s)
end

# cx:tx holds typed text or a cell binding. Excel writes both as cx:txData:
# a bound one has cx:f and a cached cx:v, a typed one only cx:v. cx:rich is
# allowed by the schema and read if present.
function _cx_tx_text(tx)::Union{Nothing,String}
    isnothing(tx) && return nothing
    rich = first_element_with_tag(tx, "rich")
    isnothing(rich) || return drawingml_text(rich)
    return child_text(first_element_with_tag(tx, "txData"), "v")
end

function _cx_tx_range(xf::XLSXFile, tx)::ChartRange
    isnothing(tx) && return nothing
    fml = child_text(first_element_with_tag(tx, "txData"), "f")
    return isnothing(fml) ? nothing : resolve_chartex_ref(xf, fml)
end

_cx_title_tx(c::ChartEx) = first_element_with_tag(first_element_with_tag(_cx_chart(c), "title"), "tx")

# ---- series and data -------------------------------------------------------

"""
    getChartSeriesCount(c::ChartEx) -> Int

Number of `cx:series` in the chart.
"""
getChartSeriesCount(c::ChartEx)::Int = length(_cx_series_nodes(c))

"""
    getSeriesLayout(c::ChartEx, i) -> Symbol

The layout of series `i` as written in its `layoutId` attribute, e.g.
`:waterfall`, `:funnel`, `:treemap`, `:clusteredColumn`, `:paretoLine`.

See also [`chartType`](@ref), which describes the chart as a whole.
"""
function getSeriesLayout(c::ChartEx, i::Integer)::Symbol
    lid = get_attr(_cx_series_node(c, i), "layoutId")
    isempty(lid) && throw(XLSXError("Series $i of chart `$(c.name)` has no layoutId."))
    return Symbol(lid)
end

function _cx_dimension(xf::XLSXFile, dim::XML.Node)::ChartExDimension
    kind = localname(dim) == "numDim" ? :num : :str
    fml  = child_text(dim, "f")
    rng  = isnothing(fml) ? nothing : resolve_chartex_ref(xf, fml)
    return ChartExDimension(kind, Symbol(get_attr(dim, "type")), fml, rng)
end

"""
    getChartDataBlocks(c::ChartEx) -> Vector{ChartExData}

Every `cx:data` block in the chart, in document order, with its dimensions.
"""
function getChartDataBlocks(c::ChartEx)::Vector{ChartExData}
    out = ChartExData[]
    chartdata = first_element_with_tag(_cx_root(c), "chartData")
    isnothing(chartdata) && return out
    for d in elements_with_tag(chartdata, "data")
        dims = [_cx_dimension(c.package, dim) for dim in XML.eachelement(d)
                if localname(dim) in ("numDim", "strDim")]
        push!(out, ChartExData(parse(Int, get_attr(d, "id")), dims))
    end
    return out
end

"""
    getSeriesData(c::ChartEx, i) -> Union{Nothing, ChartExData}

The data block series `i` draws from, found through its `cx:dataId`.
`nothing` if the series names no data block, as for a Pareto line.
"""
function getSeriesData(c::ChartEx, i::Integer)::Union{Nothing,ChartExData}
    idnode = first_element_with_tag(_cx_series_node(c, i), "dataId")
    isnothing(idnode) && return nothing
    id = parse(Int, get_attr(idnode, "val"))
    blocks = getChartDataBlocks(c)
    k = findfirst(b -> b.id == id, blocks)
    isnothing(k) &&
        throw(XLSXError("Series $i of chart `$(c.name)` refers to data block $id, which does not exist."))
    return blocks[k]
end

"""
    getSeriesOwner(c::ChartEx, i) -> Union{Nothing, Int}

For a series owned by another, such as a Pareto line, the 1-based index of the
owning series; `nothing` otherwise. Read from `ownerIdx`, taken to be the
owner's 0-based position among the chart's `cx:series`.
"""
function getSeriesOwner(c::ChartEx, i::Integer)::Union{Nothing,Int}
    s = _cx_attr(_cx_series_node(c, i), "ownerIdx")
    return isnothing(s) ? nothing : parse(Int, s) + 1
end

"""
    getSeriesHidden(c::ChartEx, i) -> Union{Nothing, Bool}

The series' `hidden` attribute. `nothing` if not written.
"""
getSeriesHidden(c::ChartEx, i::Integer) = _cx_bool(_cx_series_node(c, i), "hidden")

"""
    getSeriesAxisIds(c::ChartEx, i) -> Vector{Int}

The `id`s of the `cx:axis` elements series `i` is plotted against, as written.
These are axis ids, not positions. Empty if the series names none.
"""
getSeriesAxisIds(c::ChartEx, i::Integer)::Vector{Int} =
    [parse(Int, get_attr(n, "val")) for n in elements_with_tag(_cx_series_node(c, i), "axisId")]

# ---- names and titles ------------------------------------------------------

"""
    getSeriesName(c::ChartEx, i) -> Union{Nothing, String}

The name of series `i`, typed or the cached value of a cell binding.
`nothing` if the series has no name of its own.
"""
getSeriesName(c::ChartEx, i::Integer) =
    _cx_tx_text(first_element_with_tag(_cx_series_node(c, i), "tx"))

"""
    getSeriesNameRange(c::ChartEx, i) -> ChartRange

The cell a series name is bound to. `nothing` if the name is typed or absent.
"""
getSeriesNameRange(c::ChartEx, i::Integer) =
    _cx_tx_range(c.package, first_element_with_tag(_cx_series_node(c, i), "tx"))

"""
    getChartTitleRange(c::ChartEx) -> ChartRange

The cell the chart title is bound to. `nothing` if the title is typed or absent.
"""
getChartTitleRange(c::ChartEx) = _cx_tx_range(c.package, _cx_title_tx(c))

getChartTitleTextProps(c::ChartEx) =
    parse_drawing_text(_wb(c), first_element_with_tag(_cx_chart(c), "title"))

getLegendTextProps(c::ChartEx) =
    parse_drawing_text(_wb(c), first_element_with_tag(_cx_chart(c), "legend"))

# ---- layout properties -----------------------------------------------------

"""
    getSeriesSubtotals(c::ChartEx, i) -> Union{Nothing, Vector{Int}}

The 1-based points of a waterfall series drawn as totals. An empty vector means
`cx:subtotals` is present with no points; `nothing` means it is absent.
"""
function getSeriesSubtotals(c::ChartEx, i::Integer)::Union{Nothing,Vector{Int}}
    st = first_element_with_tag(_cx_layoutpr(c, i), "subtotals")
    isnothing(st) && return nothing
    return [parse(Int, get_attr(n, "val")) + 1 for n in elements_with_tag(st, "idx")]
end

"""
    getSeriesBinning(c::ChartEx, i) -> Union{Nothing, ChartExBinning}

Histogram binning for series `i`; `nothing` if the series has no `cx:binning`.
"""
function getSeriesBinning(c::ChartEx, i::Integer)::Union{Nothing,ChartExBinning}
    b = first_element_with_tag(_cx_layoutpr(c, i), "binning")
    isnothing(b) && return nothing
    return ChartExBinning(
        _cx_symbol(b, "intervalClosed"),
        _cx_num_or_auto(b, "underflow", Float64),
        _cx_num_or_auto(b, "overflow", Float64),
        _cx_num_or_auto(first_element_with_tag(b, "binSize"), "val", Float64),
        _cx_num_or_auto(first_element_with_tag(b, "binCount"), "val", Int))
end

"""
    getSeriesAggregation(c::ChartEx, i) -> Bool

Whether series `i` aggregates repeated categories (`cx:aggregation`). The
element carries no value; its presence is the setting.
"""
getSeriesAggregation(c::ChartEx, i::Integer)::Bool =
    !isnothing(first_element_with_tag(_cx_layoutpr(c, i), "aggregation"))

"""
    getSeriesQuartileMethod(c::ChartEx, i) -> Union{Nothing, Symbol}

Box & whisker quartile calculation, `:inclusive` or `:exclusive`.
"""
getSeriesQuartileMethod(c::ChartEx, i::Integer) =
    _cx_symbol(first_element_with_tag(_cx_layoutpr(c, i), "statistics"), "quartileMethod")

"""
    getSeriesParentLabelLayout(c::ChartEx, i) -> Union{Nothing, Symbol}

Treemap parent label placement, e.g. `:overlapping`, `:banner`, `:none`.
"""
getSeriesParentLabelLayout(c::ChartEx, i::Integer) =
    _cx_symbol(first_element_with_tag(_cx_layoutpr(c, i), "parentLabelLayout"), "val")

const _CX_LAYOUT_FLAGS = (:meanLine, :meanMarker, :nonoutliers, :outliers, :connectorLines)

"""
    getSeriesLayoutFlag(c::ChartEx, i, flag) -> Union{Nothing, Bool}

One attribute of `cx:layoutPr/cx:visibility`: `:meanLine`, `:meanMarker`,
`:nonoutliers`, `:outliers` (box & whisker) or `:connectorLines` (waterfall).
`nothing` if not written.
"""
function getSeriesLayoutFlag(c::ChartEx, i::Integer, flag::Symbol)::Union{Nothing,Bool}
    flag in _CX_LAYOUT_FLAGS ||
        throw(XLSXError("Unknown layout flag `$flag`. Expected one of $(join(_CX_LAYOUT_FLAGS, ", "))."))
    return _cx_bool(first_element_with_tag(_cx_layoutpr(c, i), "visibility"), String(flag))
end

# ---- data labels -----------------------------------------------------------

const _CX_LABEL_FLAGS = (:seriesName, :categoryName, :value)

"""
    getLabelFlag(c::ChartEx, i, flag) -> Union{Nothing, Bool}

Whether series `i`'s data labels show `:seriesName`, `:categoryName` or
`:value`, from `cx:dataLabels/cx:visibility`. `nothing` if not written.
"""
function getLabelFlag(c::ChartEx, i::Integer, flag::Symbol)::Union{Nothing,Bool}
    flag in _CX_LABEL_FLAGS ||
        throw(XLSXError("Unknown label flag `$flag`. Expected one of $(join(_CX_LABEL_FLAGS, ", "))."))
    return _cx_bool(first_element_with_tag(_cx_datalabels(c, i), "visibility"), String(flag))
end

"""
    getLabelPosition(c::ChartEx, i) -> Union{Nothing, Symbol}

Data label position for series `i`, e.g. `:outEnd`, `:inEnd`, `:ctr`.
"""
getLabelPosition(c::ChartEx, i::Integer) = _cx_symbol(_cx_datalabels(c, i), "pos")

# ---- formatting cascades ---------------------------------------------------

# The cx:dataPt for 1-based `point` of series `i`, or `nothing` if Excel never
# formatted that point. cx:dataPt idx is 0-based.
function _cx_datapt(c::ChartEx, i::Integer, point::Integer)
    point >= 1 || throw(XLSXError("Data point positions start at 1; asked for $point."))
    for dp in elements_with_tag(_cx_series_node(c, i), "dataPt")
        tryparse(Int, get_attr(dp, "idx")) == point - 1 && return dp
    end
    return nothing
end

# spPr cascade: data point, then series. As for c:, nothing above the series
# describes a series' graphic; the rest comes from the style part.
function _cx_shape_chain(c::ChartEx, i::Integer; point::Union{Nothing,Integer}=nothing)
    chain = FormatSite[]
    if !isnothing(point)
        dp = _cx_datapt(c, i, point)
        isnothing(dp) || push!(chain, _site_path(:point, :shape, dp, "spPr"))
    end
    push!(chain, _site_path(:series, :shape, _cx_series_node(c, i), "spPr"))
    return chain
end

# txPr cascade for data labels: series, then chart space. cx has no chart groups.
_cx_text_chain(c::ChartEx, i::Integer) = FormatSite[
    _site_path(:series,     :text, _cx_series_node(c, i), "dataLabels", "txPr"),
    _site_path(:chartspace, :text, _cx_root(c),           "txPr"),
]

"""
    getSeriesFill(c::ChartEx, i::Integer; point=nothing) -> Effective{DrawingFill}

Resolve the fill of series `i`, or of one data point of it, walking
`cx:dataPt/cx:spPr` then `cx:series/cx:spPr`.
"""
getSeriesFill(c::ChartEx, i::Integer; point::Union{Nothing,Integer}=nothing) =
    _walk_fill(_wb(c), _cx_shape_chain(c, i; point))

"""
    getLabelTextProp(c::ChartEx, i::Integer, field::Symbol) -> Effective

Resolve one field of the data-label text formatting for series `i`, walking
`cx:dataLabels/cx:txPr` then `cx:chartSpace/cx:txPr`. See the `Chart` method
for how fields inherit.
"""
getLabelTextProp(c::ChartEx, i::Integer, field::Symbol) =
    _resolve_text_field(_wb(c), _cx_text_chain(c, i), field)