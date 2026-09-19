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
    getSeriesLine(c::ChartEx, i::Integer; point=nothing) -> Effective{DrawingLine}

Resolve the outline of series `i`, or of one data point of it, walking
`cx:dataPt/cx:spPr` then `cx:series/cx:spPr`.
"""
getSeriesLine(c::ChartEx, i::Integer; point::Union{Nothing,Integer} = nothing) =
    _walk_line(_wb(c), _cx_shape_chain(c, i; point))

"""
    getLabelTextProp(c::ChartEx, i::Integer, field::Symbol) -> Effective

Resolve one field of the data-label text formatting for series `i`, walking
`cx:dataLabels/cx:txPr` then `cx:chartSpace/cx:txPr`. See the `Chart` method
for how fields inherit.
"""
getLabelTextProp(c::ChartEx, i::Integer, field::Symbol) =
    _resolve_text_field(_wb(c), _cx_text_chain(c, i), field)

    # ---- write path -------------------------------------------------------------
#
# A ChartEx is durable, so setters write the part and return `c` itself, which
# remains valid. Handles to nodes (anything holding `raw`) are still invalidated,
# as for c:.

const _CX_ROOT_KEY = (NS_CX, "chartSpace")

# rebuild_path steps from the root to series `i`, matched by node identity.
function _cx_series_path(c::ChartEx, i::Integer)
    ser = _cx_series_node(c, i)
    return [(NS_CX, "chart")          => "chart",
            (NS_CX, "plotArea")       => "plotArea",
            (NS_CX, "plotAreaRegion") => "plotAreaRegion",
            (NS_CX, "series")         => ("series", n -> n === ser)]
end

# Apply `f(ser, pfx)` to series `i` and write the part back once. A throw inside
# `f` leaves the part untouched.
function _cx_edit_series!(c::ChartEx, i::Integer, f)
    root = _cx_root(c)
    pfx  = ns_prefixes(root)
    new  = rebuild_path(root, _cx_series_path(c, i), ser -> f(ser, pfx);
                        prefixes = pfx, parent_key = _CX_ROOT_KEY)
    set_chart_root!(c, new)
    return c
end

_cx_idx(n) = tryparse(Int, get_attr(n, "idx"))

# Apply `f` to the cx:dataPt for 1-based `point`, creating it if absent. A new
# one is inserted in idx order among its siblings: before the first dataPt with a
# larger idx, or in schema position (after the last dataPt) if there is none.
function _cx_with_datapt(ser::XML.Node, point::Integer, pfx, f)
    idx  = point - 1
    kids = isnothing(ser.children) ? XML.Node[] : ser.children
    j = findfirst(k -> localname(k) == "dataPt" && _cx_idx(k) == idx, kids)
    if isnothing(j)
        fresh = XML.Element(prefixed_tag(pfx[NS_CX], "dataPt"); idx = string(idx))
        pts = findall(k -> localname(k) == "dataPt", kids)
        if isempty(pts)
            ser = insert_child(ser, (NS_CX, "series"), fresh)      # first one: schema position
        else
            # before the first dataPt with a larger idx, else after the last dataPt
            k = findfirst(p -> something(_cx_idx(kids[p]), -1) > idx, pts)
            at = isnothing(k) ? last(pts) + 1 : pts[k]
            new_kids = Vector{eltype(kids)}(undef, 0)
            append!(new_kids, kids[1:at-1]); push!(new_kids, fresh); append!(new_kids, kids[at:end])
            ser = _with_children(ser, new_kids)
        end
        j = findfirst(n -> n === fresh, ser.children)
    end
    dp = ser.children[j]
    return replace_child(ser, dp, f(dp))
end

# Apply `f(spPr, pfx)` to the spPr of series `i`, or of one of its points.
# With `create = false`, a missing spPr (or dataPt) means there is nothing to
# change, and nothing is created: removing a fill that was never written must
# not leave an empty <cx:spPr/> behind.
function _cx_set_shape!(c::ChartEx, i::Integer, point, f; create::Bool = true)
    isnothing(point) || point >= 1 ||
        throw(XLSXError("Data point positions start at 1; asked for $point."))
    _cx_edit_series!(c, i, (ser, pfx) -> begin
        edit(node, key) =
            (!create && isnothing(first_element_with_tag(node, "spPr"))) ? node :
            rebuild_path(node, [(NS_A, "spPr") => "spPr"], sp -> f(sp, pfx);
                         prefixes = pfx, parent_key = key)
        isnothing(point) && return edit(ser, (NS_CX, "series"))
        (!create && isnothing(_cx_datapt_in(ser, point))) && return ser
        return _cx_with_datapt(ser, point, pfx, dp -> edit(dp, (NS_CX, "dataPt")))
    end)
end

_cx_datapt_in(ser, point) =
    findfirst(k -> localname(k) == "dataPt" && _cx_idx(k) == point - 1,
              isnothing(ser.children) ? XML.Node[] : ser.children)

"""
    setSeriesFill(c::ChartEx, i, color; point=nothing) -> ChartEx

Set the fill of series `i`, or of one data point of it. `color` may be a colour
string or Symbol, a `Colors.Colorant`, a [`SchemeColor`](@ref), `:none` for an
explicit `<a:noFill/>`, or `:inherit` to remove the fill so the chart style
applies.

Setting a point's fill creates its `cx:dataPt` if Excel never formatted that
point. `:inherit` never creates anything.

Returns `c`, which remains valid.
"""
setSeriesFill(c::ChartEx, i::Integer, color::Union{AbstractString,Colors.Colorant,SchemeColor};
              point::Union{Nothing,Integer} = nothing) =
    _cx_set_shape!(c, i, point, (sp, pfx) -> _sp_with_fill(sp, (NS_A, "spPr"), color, pfx))

function setSeriesFill(c::ChartEx, i::Integer, what::Symbol;
                       point::Union{Nothing,Integer} = nothing)
    what === :inherit &&
        return _cx_set_shape!(c, i, point, (sp, pfx) -> _sp_with_fill(sp, (NS_A, "spPr"), :inherit, pfx);
                              create = false)
    what === :none &&
        return _cx_set_shape!(c, i, point, (sp, pfx) -> _sp_with_fill(sp, (NS_A, "spPr"), :none, pfx))
    return setSeriesFill(c, i, String(what); point)
end

# Apply `f(ln, pfx)` to the a:ln of series `i`, or of one of its points,
# creating the a:ln and its cx:spPr if absent. One rebuild; a throw inside `f`
# leaves the part untouched.
_cx_set_line!(c::ChartEx, i::Integer, point, f; create::Bool = true) =
    _cx_set_shape!(c, i, point,
                   (sp, pfx) -> _sp_with_line(sp, (NS_A, "spPr"), ln -> f(ln, pfx), pfx);
                   create)

"""
    setSeriesLine(c::ChartEx, i; point=nothing, color, width, dash, cap, compound, join, miterLimit) -> ChartEx
    setSeriesLine(c::ChartEx, i, :none; point=nothing)
    setSeriesLine(c::ChartEx, i, :inherit; point=nothing)

Set several outline properties of series `i`, or of one data point of it, in one
rebuild of the chart part. A keyword left unspecified is left alone; pass
`:inherit` to remove one that is set.

The symbol form acts on the whole outline: `:none` writes an `a:ln` whose fill is
`<a:noFill/>`, and `:inherit` removes the `a:ln` so the chart style supplies it.

Returns `c`, which remains valid.
"""
function setSeriesLine(c::ChartEx, i::Integer; point::Union{Nothing,Integer} = nothing,
                       color = nothing, width = nothing, dash = nothing,
                       cap = nothing, compound = nothing,
                       join = nothing, miterLimit = nothing)
    all(isnothing, (color, width, dash, cap, compound, join, miterLimit)) && return c
    return _cx_set_line!(c, i, point, (ln, pfx) ->
        _ln_with(ln, pfx; color, width, dash, cap, compound, join, miterLimit))
end

function setSeriesLine(c::ChartEx, i::Integer, what::Symbol;
                       point::Union{Nothing,Integer} = nothing)
    what === :inherit &&
        return _cx_set_shape!(c, i, point, (sp, _) -> remove_child(sp, "ln"); create = false)
    what === :none && return setSeriesLineColor(c, i, :none; point)
    throw(XLSXError("`$what` is not a line instruction; use `:none` or `:inherit`."))
end

"""
    setSeriesLineColor(c::ChartEx, i, color; point=nothing) -> ChartEx

Set the colour of series `i`'s outline, or of one data point of it. `color` takes
the same values as [`setSeriesFill`](@ref): `:none` writes an outline with
`<a:noFill/>`, `:inherit` removes the colour and leaves the rest of the `a:ln`.
"""
setSeriesLineColor(c::ChartEx, i::Integer, color; point::Union{Nothing,Integer} = nothing) =
    _cx_set_line!(c, i, point, (ln, pfx) -> _ln_with_color(ln, color, pfx))

setSeriesLineWidth(c::ChartEx, i::Integer, pts; point::Union{Nothing,Integer} = nothing) =
    _cx_set_line!(c, i, point, (ln, _) -> _ln_with_width(ln, pts))

setSeriesLineDash(c::ChartEx, i::Integer, dash; point::Union{Nothing,Integer} = nothing) =
    _cx_set_line!(c, i, point, (ln, pfx) -> _ln_with_dash(ln, dash, pfx))

setSeriesLineCap(c::ChartEx, i::Integer, cap; point::Union{Nothing,Integer} = nothing) =
    _cx_set_line!(c, i, point, (ln, _) -> _ln_with_cap(ln, cap))

setSeriesLineCompound(c::ChartEx, i::Integer, cmpd; point::Union{Nothing,Integer} = nothing) =
    _cx_set_line!(c, i, point, (ln, _) -> _ln_with_compound(ln, cmpd))

setSeriesLineJoin(c::ChartEx, i::Integer, join; point::Union{Nothing,Integer} = nothing) =
    _cx_set_line!(c, i, point, (ln, pfx) -> _ln_with_join(ln, join, pfx))

setSeriesLineMiterLimit(c::ChartEx, i::Integer, lim; point::Union{Nothing,Integer} = nothing) =
    _cx_set_line!(c, i, point, (ln, _) -> _ln_with_miter_limit(ln, lim))

    # A created cx:dataLabels showing no flags would show labels with Excel's
# defaults, as c:dLbls does, so all three are written off when creating one.
# Unlike c:, they are attributes on a single cx:visibility child.
const _CX_LABEL_FLAG_NAMES = ("seriesName", "categoryName", "value")

function _cx_visibility_off(lbl::XML.Node, pfx::Dict{String,String})
    isnothing(first_element_with_tag(lbl, "visibility")) || return lbl
    vis = XML.Element(prefixed_tag(pfx[NS_CX], "visibility");
                      seriesName = "0", categoryName = "0", value = "0")
    return insert_child(lbl, (NS_CX, "dataLabels"), vis)
end

"""
    setLabelTextProp(c::ChartEx, i, field, value) -> ChartEx

Set one field of the data-label text formatting for series `i`, writing it on the
series' `cx:dataLabels/cx:txPr`. `field` is a field of `DrawingRunProps`; `:fill`
and `:line` are composite and take a colour or a named tuple of line properties.
Pass `:inherit` to remove a field so the cascade resolves it.

Creating a `cx:dataLabels` to hold formatting writes `cx:visibility` with all
three flags off, since one that names no flags shows labels with Excel's
defaults rather than none.

Returns `c`, which remains valid.
"""
function setLabelTextProp(c::ChartEx, i::Integer, field::Symbol, value)
    field in _RUN_PROP_FIELDS || throw(XLSXError(
        "`$field` is not a resolvable text property. Valid fields: " *
        join(_RUN_PROP_FIELDS, ", ") * "."))
    return _cx_edit_series!(c, i, (ser, pfx) ->
        rebuild_path(ser, [(NS_CX, "dataLabels") => "dataLabels"],
                     lbl -> _cx_label_transform(lbl, field, value, pfx);
                     prefixes = pfx, parent_key = (NS_CX, "series")))
end

# Apply `field` to the cx:dataLabels' cx:txPr, creating the text body and the
# visibility flags if this is the first formatting written here.
function _cx_label_transform(lbl::XML.Node, field::Symbol, value, pfx::Dict{String,String})
    lbl = _cx_visibility_off(lbl, pfx)
    tx  = first_element_with_tag(lbl, "txPr")
    if isnothing(tx)
        lbl = insert_child(lbl, (NS_CX, "dataLabels"), _new_text_body("txPr", pfx, NS_CX))
        tx  = first_element_with_tag(lbl, "txPr")
    end
    return replace_child(lbl, tx, _text_with_run_prop(tx, field, value, pfx))
end

# ---- chart-level text -------------------------------------------------------

# Apply `f(node, pfx)` at the end of `steps` from the part root, and write back.
function _cx_edit_at!(c::ChartEx, steps, f)
    root = _cx_root(c)
    pfx  = ns_prefixes(root)
    new  = rebuild_path(root, steps, n -> f(n, pfx);
                        prefixes = pfx, parent_key = _CX_ROOT_KEY)
    set_chart_root!(c, new)
    return c
end

"""
    setChartTitleTextProp(c::ChartEx, field, value) -> ChartEx

Set one field of the chart title's text formatting, on `cx:title/cx:txPr`.
Creates the title if the chart has none. Returns `c`, which remains valid.
"""
setChartTitleTextProp(c::ChartEx, field::Symbol, value) =
    _cx_edit_at!(c, [(NS_CX, "chart") => "chart",
                     (NS_CX, "title") => "title"],
                 (t, pfx) -> _both_with_run_prop(t, (NS_CX, "title"), field, value, pfx; ns = NS_CX))

"""
    setLegendTextProp(c::ChartEx, field, value) -> ChartEx

Set one field of the legend's text formatting, on `cx:legend/cx:txPr`.
Creates the legend if the chart has none. Returns `c`, which remains valid.
"""
setLegendTextProp(c::ChartEx, field::Symbol, value) =
    _cx_edit_at!(c, [(NS_CX, "chart")  => "chart",
                     (NS_CX, "legend") => "legend"],
                 (lg, pfx) -> _txpr_with_run_prop(lg, (NS_CX, "legend"), field, value, pfx; ns = NS_CX))

"""
    getChartAxisIds(c::ChartEx) -> Vector{Int}

The `id` of each `cx:axis`, in document order. These are the values
[`getSeriesAxisIds`](@ref) refers to.
"""
getChartAxisIds(c::ChartEx)::Vector{Int} =
    [parse(Int, get_attr(a, "id")) for a in
     elements_with_tag(first_element_with_tag(_cx_chart(c), "plotArea"), "axis")]

# Steps to the cx:axis whose id attribute is `id`.
function _cx_axis_path(c::ChartEx, id::Integer)
    ids = getChartAxisIds(c)
    id in ids || throw(XLSXError(
        "Chart `$(c.name)` has no axis with id $id. Axis ids: $(join(ids, ", "))."))
    return [(NS_CX, "chart")    => "chart",
            (NS_CX, "plotArea") => "plotArea",
            (NS_CX, "axis")     => ("axis", n -> tryparse(Int, get_attr(n, "id")) == id)]
end

"""
    setAxisTitleTextProp(c::ChartEx, id::Integer, field, value) -> ChartEx

Set one field of an axis title's text formatting, on `cx:axis/cx:title/cx:txPr`.
`id` is the axis' `id` attribute, as [`getChartAxisIds`](@ref) reports it, not a
position. Creates the title if the axis has none. Returns `c`, which remains
valid.
"""
setAxisTitleTextProp(c::ChartEx, id::Integer, field::Symbol, value) =
    _cx_edit_at!(c, [_cx_axis_path(c, id)...; (NS_CX, "title") => "title"],
                 (t, pfx) -> _both_with_run_prop(t, (NS_CX, "title"), field, value, pfx; ns = NS_CX))

"""
    getAxisTitleTextProps(c::ChartEx, id::Integer) -> Union{Nothing, DrawingText}

The text body of the axis title, or `nothing` if the axis has no title.
"""
getAxisTitleTextProps(c::ChartEx, id::Integer) =
    parse_drawing_text(_wb(c), first_element_with_tag(_cx_axis_node(c, id), "title"))

function _cx_axis_node(c::ChartEx, id::Integer)
    pa = first_element_with_tag(_cx_chart(c), "plotArea")
    for a in elements_with_tag(pa, "axis")
        tryparse(Int, get_attr(a, "id")) == id && return a
    end
    throw(XLSXError("Chart `$(c.name)` has no axis with id $id."))
end

# ---- text content -----------------------------------------------------------

# Replace cx:tx with a typed cx:txData/cx:v holding `text`. Any cx:f is dropped:
# typing text unbinds the element from its cell.
function _cx_tx_typed(tx::XML.Node, text::AbstractString, pfx::Dict{String,String})
    v  = XML.Element(prefixed_tag(pfx[NS_CX], "v"), XML.Text(text))
    td = XML.Element(prefixed_tag(pfx[NS_CX], "txData"), v)
    return _with_children(tx, [td])
end

# Excel writes a title's text twice: in cx:txData/cx:v and as runs in the sibling
# cx:txPr, which is what it renders. Collapse the body to a single run carrying
# `text`, keeping the first run's formatting where there was one, so the words
# change and the appearance does not.
function _cx_txpr_retext(el::XML.Node, text::AbstractString, pfx::Dict{String,String})
    tx = first_element_with_tag(el, "txPr")
    isnothing(tx) && return el
    isnothing(tx.children) && return el
    kids = map(tx.children) do p
        localname(p) != "p" && return p
        old  = first_element_with_tag(p, "r")
        rpr  = isnothing(old) ? nothing : first_element_with_tag(old, "rPr")
        run  = XML.Element(prefixed_tag(pfx[NS_A], "r"),
                           filter(!isnothing, [rpr,
                               XML.Element(prefixed_tag(pfx[NS_A], "t"), XML.Text(text))])...)
        keep = filter(n -> localname(n) in ("pPr", "endParaRPr"),
                      isnothing(p.children) ? XML.Node[] : p.children)
        ppr  = filter(n -> localname(n) == "pPr", keep)
        epr  = filter(n -> localname(n) == "endParaRPr", keep)
        return _with_children(p, [ppr; run; epr])
    end
    return replace_child(el, tx, _with_children(tx, kids))
end

"""
    setChartTitleText(c::ChartEx, text::AbstractString) -> ChartEx

Set the chart title to `text`, typed rather than bound to a cell. Creates the
title if the chart has none, and drops any `cx:f` binding it had.

Excel stores a typed title both as `cx:tx/cx:txData/cx:v` and as runs in the
title's `cx:txPr`, and renders the latter, so both are updated. Run formatting
is preserved; use [`setChartTitleTextProp`](@ref) to change it.

Returns `c`, which remains valid.
"""
setChartTitleText(c::ChartEx, text::AbstractString) =
    _cx_edit_at!(c, [(NS_CX, "chart") => "chart",
                     (NS_CX, "title") => "title"],
                 (t, pfx) -> begin
                     t = rebuild_path(t, [(NS_CX, "tx") => "tx"],
                                      tx -> _cx_tx_typed(tx, text, pfx);
                                      prefixes = pfx, parent_key = (NS_CX, "title"))
                     return _cx_txpr_retext(t, text, pfx)
                 end)

"""
    setSeriesName(c::ChartEx, i, name::AbstractString) -> ChartEx

Set the name of series `i` to `name`, typed rather than bound to a cell.
Creates the series' `cx:tx` if it has none, and drops any `cx:f` binding.

Returns `c`, which remains valid.
"""
setSeriesName(c::ChartEx, i::Integer, name::AbstractString) =
    _cx_edit_series!(c, i, (ser, pfx) ->
        rebuild_path(ser, [(NS_CX, "tx") => "tx"],
                     tx -> _cx_tx_typed(tx, name, pfx);
                     prefixes = pfx, parent_key = (NS_CX, "series")))
 
# ---- layout properties (write) ----------------------------------------------

# Apply `f(layoutPr, pfx)` to series `i`'s cx:layoutPr, creating it if absent.
_cx_set_layoutpr!(c::ChartEx, i::Integer, f) =
    _cx_edit_series!(c, i, (ser, pfx) ->
        rebuild_path(ser, [(NS_CX, "layoutPr") => "layoutPr"], lp -> f(lp, pfx);
                     prefixes = pfx, parent_key = (NS_CX, "series")))

"""
    setSeriesSubtotals(c::ChartEx, i, points) -> ChartEx

Mark the given points of a waterfall series as totals, replacing whatever was
marked before. `points` counts from 1, in any order; an empty collection writes
`<cx:subtotals/>`, which is what Excel writes for a waterfall with no totals.

Pass `:inherit` to remove `cx:subtotals` entirely.

Returns `c`, which remains valid.
"""
function setSeriesSubtotals(c::ChartEx, i::Integer, points)
    if points === :inherit
        return _cx_set_layoutpr!(c, i, (lp, _) -> remove_child(lp, "subtotals"))
    end
    idx = sort(unique(Int[p for p in points]))
    isempty(idx) || first(idx) >= 1 ||
        throw(XLSXError("Data point positions start at 1; asked for $(first(idx))."))
    return _cx_set_layoutpr!(c, i, (lp, pfx) -> begin
        st = XML.Element(prefixed_tag(pfx[NS_CX], "subtotals"),
                         (XML.Element(prefixed_tag(pfx[NS_CX], "idx"); val = string(p - 1))
                          for p in idx)...)
        return insert_child(lp, (NS_CX, "layoutPr"), st)
    end)
end

# Set or remove one attribute on a single-purpose cx:layoutPr child, creating the
# child if needed. `:inherit` removes the attribute, and the child with it when
# that leaves it empty, since these elements exist only to carry their attributes.
function _cx_set_layout_attr!(c::ChartEx, i::Integer, tag::AbstractString,
                              attr::AbstractString, value)
    return _cx_set_layoutpr!(c, i, (lp, pfx) -> begin
        if value === :inherit
            return remove_child(lp, tag)
        end
        node = something(first_element_with_tag(lp, tag),
                         XML.Element(prefixed_tag(pfx[NS_CX], tag)))
        return insert_child(lp, (NS_CX, "layoutPr"), with_attribute(node, attr, value))
    end)
end

const _CX_QUARTILE_METHODS = (:inclusive, :exclusive)

"""
    setSeriesQuartileMethod(c::ChartEx, i, method) -> ChartEx

Set the quartile calculation of a box & whisker series to `:inclusive` or
`:exclusive`, or `:inherit` to remove `cx:statistics`.
"""
function setSeriesQuartileMethod(c::ChartEx, i::Integer, method::Symbol)
    method === :inherit || method in _CX_QUARTILE_METHODS ||
        throw(XLSXError("`$method` is not a quartile method. Use " *
                        join(_CX_QUARTILE_METHODS, " or ") * "."))
    return _cx_set_layout_attr!(c, i, "statistics", "quartileMethod",
                                method === :inherit ? :inherit : String(method))
end

const _CX_PARENT_LABEL_LAYOUTS = (:none, :banner, :overlapping)

"""
    setSeriesParentLabelLayout(c::ChartEx, i, layout) -> ChartEx

Set a treemap series' parent label placement to `:none`, `:banner` or
`:overlapping`, or `:inherit` to remove `cx:parentLabelLayout`.
"""
function setSeriesParentLabelLayout(c::ChartEx, i::Integer, layout::Symbol)
    layout === :inherit || layout in _CX_PARENT_LABEL_LAYOUTS ||
        throw(XLSXError("`$layout` is not a parent label layout. Valid values: " *
                        join(_CX_PARENT_LABEL_LAYOUTS, ", ") * "."))
    return _cx_set_layout_attr!(c, i, "parentLabelLayout", "val",
                                layout === :inherit ? :inherit : String(layout))
end

"""
    setSeriesLayoutFlag(c::ChartEx, i, flag, value) -> ChartEx

Set one attribute of `cx:layoutPr/cx:visibility`: `:meanLine`, `:meanMarker`,
`:nonoutliers`, `:outliers` (box & whisker) or `:connectorLines` (waterfall).
`value` is a `Bool`, or `:inherit` to remove the attribute.
"""
function setSeriesLayoutFlag(c::ChartEx, i::Integer, flag::Symbol, value)
    flag in _CX_LAYOUT_FLAGS ||
        throw(XLSXError("Unknown layout flag `$flag`. Expected one of $(join(_CX_LAYOUT_FLAGS, ", "))."))
    value isa Bool || value === :inherit ||
        throw(XLSXError("A layout flag takes a Bool or `:inherit`; got `$value`."))
    return _cx_set_layoutpr!(c, i, (lp, pfx) -> begin
        vis = something(first_element_with_tag(lp, "visibility"),
                        XML.Element(prefixed_tag(pfx[NS_CX], "visibility")))
        vis = with_attribute(vis, String(flag), value === :inherit ? nothing : (value ? "1" : "0"))
        return isnothing(vis.attributes) || isempty(vis.attributes) ?
               remove_child(lp, "visibility") :
               insert_child(lp, (NS_CX, "layoutPr"), vis)
    end)
end

"""
    setSeriesAggregation(c::ChartEx, i, on::Bool) -> ChartEx

Turn category aggregation on or off for series `i`. `cx:aggregation` carries no
value, so its presence is the setting. It excludes `cx:binning`, which is
removed when this is turned on.
"""
setSeriesAggregation(c::ChartEx, i::Integer, on::Bool) =
    _cx_set_layoutpr!(c, i, (lp, pfx) -> on ?
        insert_child(remove_choice(lp, (NS_CX, "layoutPr"), AGGREGATION_GROUP),
                     (NS_CX, "layoutPr"),
                     XML.Element(prefixed_tag(pfx[NS_CX], "aggregation"))) :
        remove_child(lp, "aggregation"))

"""
    setSeriesBinning(c::ChartEx, i; intervalClosed, underflow, overflow, binSize, binCount) -> ChartEx

Set histogram binning for series `i` in one rebuild. A keyword left unspecified
is left alone; pass `:inherit` to remove one that is set.

`intervalClosed` is `:r` or `:l`. `underflow` and `overflow` take a number or
`:auto`. `binSize` and `binCount` are alternatives: setting one removes the
other. `cx:binning` excludes `cx:aggregation`, which is removed if present.
"""
function setSeriesBinning(c::ChartEx, i::Integer;
                          intervalClosed = nothing, underflow = nothing, overflow = nothing,
                          binSize = nothing, binCount = nothing)
    all(isnothing, (intervalClosed, underflow, overflow, binSize, binCount)) && return c
    isnothing(binSize) || isnothing(binCount) ||
        throw(XLSXError("`binSize` and `binCount` are alternatives; set only one."))
    isnothing(intervalClosed) || intervalClosed === :inherit || intervalClosed in (:r, :l) ||
        throw(XLSXError("`intervalClosed` is `:r` or `:l`; got `$intervalClosed`."))

    fmt(v) = v === :auto ? "auto" : string(v)

    return _cx_set_layoutpr!(c, i, (lp, pfx) -> begin
        lp = remove_child(lp, "aggregation")
        b  = something(first_element_with_tag(lp, "binning"),
                       XML.Element(prefixed_tag(pfx[NS_CX], "binning")))
        isnothing(intervalClosed) ||
            (b = with_attribute(b, "intervalClosed",
                                intervalClosed === :inherit ? nothing : String(intervalClosed)))
        isnothing(underflow) ||
            (b = with_attribute(b, "underflow", underflow === :inherit ? nothing : fmt(underflow)))
        isnothing(overflow) ||
            (b = with_attribute(b, "overflow", overflow === :inherit ? nothing : fmt(overflow)))
        for (kw, tag) in ((binSize, "binSize"), (binCount, "binCount"))
            isnothing(kw) && continue
            b = remove_choice(b, (NS_CX, "binning"), BIN_GROUP)
            kw === :inherit && continue
            b = insert_child(b, (NS_CX, "binning"),
                             XML.Element(prefixed_tag(pfx[NS_CX], tag); val = fmt(kw)))
        end
        return insert_child(lp, (NS_CX, "layoutPr"), b)
    end)
end