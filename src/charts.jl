#
# charts.jl
#
# Read chart metadata and the cached data Excel stores inside chart parts
# (JuliaData/XLSX.jl#263).
#
# Excel writes a snapshot of every series' source data into the chart part
# itself (`c:numCache` / `c:strCache` / `c:multiLvlStrCache`). That cache is what
# makes a chart render when its source is unavailable - a deleted sheet, or an
# external workbook that isn't to hand - and it is what this file exposes. It is
# never re-read from the worksheet, so it reflects the values as of the last time
# Excel saved the file.
#
# Two schemas are handled. The original `c:` schema covers the sixteen classic
# plot types and is read in full. The newer `cx:` schema (waterfall, funnel,
# treemap, sunburst, histogram, Pareto, box & whisker, region map) is read for
# discovery only: type, title and source ranges, but no cached values.
#

const REL_CHART   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
const REL_CHARTEX = "http://schemas.microsoft.com/office/2014/relationships/chartEx"

const CT_CHART   = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
const CT_CHARTEX = "application/vnd.ms-office.chartex+xml"

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

const NS_C  = "http://schemas.openxmlformats.org/drawingml/2006/chart"
const NS_CX = "http://schemas.microsoft.com/office/drawing/2014/chart"

# The <c:plotArea> children that group series. Series live one level below these.
const CHART_GROUP_TAGS = Set([
    "areaChart", "area3DChart", "lineChart", "line3DChart", "stockChart",
    "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart",
    "barChart", "bar3DChart", "ofPieChart", "surfaceChart", "surface3DChart",
    "bubbleChart",
])

# Drawing anchor children that can hold a graphic frame, directly or nested.
const SHAPE_TAGS = ("graphicFrame", "pic", "sp", "grpSp", "cxnSp", "contentPart")


# ===========================================================================
# Traversal helpers
# ===========================================================================

"""
Id => (resolved target path, relationship type) for the relationships of
`part_path` whose type is in `reltypes`. Targets are resolved against the part's
own directory, so `../charts/chart1.xml` from `xl/drawings/drawing1.xml` gives
`xl/charts/chart1.xml`.
"""
function rid_to_target(xf::XLSXFile, part_path::String, reltypes)::Dict{String,Tuple{String,String}}
    dir, fname = _split_zip_path(part_path)
    rels_path = isempty(dir) ? "_rels/$fname.rels" : "$dir/_rels/$fname.rels"
    targets = Dict{String,Tuple{String,String}}()
    haskey(xf.data, rels_path) || return targets
    for n in elements_with_tag(xml_root_element(xf.data[rels_path]), "Relationship")
        rt = get_attr(n, "Type")
        rt in reltypes && get_attr(n, "TargetMode") != "External" || continue
        id = get_attr(n, "Id")
        isempty(id) && continue
        targets[id] = (resolve_relative_target(dir, get_attr(n, "Target")), rt)
    end
    return targets
end

"Child elements of `node` whose local name is in `tags`."
function elements_with_tags(node::XML.Node, tags)::Vector{XML.Node}
    out = XML.Node[]
    for child in XML.eachelement(node)
        localname(child) in tags && push!(out, child)
    end
    return out
end

"""
The effective shape elements of a drawing anchor, transparently unwrapping any
`mc:AlternateContent`.

Charts using the `cx:` schema are always written inside `mc:AlternateContent`,
with the real graphic frame under `mc:Choice` and a static picture under
`mc:Fallback`. Without this unwrapping they are invisible to discovery. Newer
picture effects and slicers use the same wrapper, so this is deliberately
generic rather than a `chartEx` special case.
"""
function effective_shapes(anchor::XML.Node)::Vector{XML.Node}
    out = XML.Node[]
    for child in XML.eachelement(anchor)
        name = localname(child)
        if name == "AlternateContent"
            append!(out, resolve_alternate_content(child))
        elseif name in SHAPE_TAGS
            push!(out, child)
        end
    end
    return out
end

function resolve_alternate_content(ac::XML.Node)::Vector{XML.Node}
    fallback = nothing
    for branch in XML.eachelement(ac)
        n = localname(branch)
        if n == "Choice"
            # The first Choice unconditionally: branches differ in fidelity, not
            # in which chart they reference, and discovery only needs to see the
            # frame. Checking `Requires` matters only once we render.
            return elements_with_tags(branch, SHAPE_TAGS)
        elseif n == "Fallback" && isnothing(fallback)
            fallback = branch
        end
    end
    return isnothing(fallback) ? XML.Node[] : elements_with_tags(fallback, SHAPE_TAGS)
end

# The chart reference sits at a fixed depth below a graphic frame:
#   <xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="..."/>
# `cx:chart` sits in exactly the same place, so one walk covers both; the
# relationship type, not the element, tells the two schemas apart.
function frame_chart_element(shape::XML.Node)::Union{Nothing,XML.Node}
    localname(shape) == "graphicFrame" || return nothing
    graphic = first_element_with_tag(shape, "graphic")
    graphicdata = first_element_with_tag(graphic, "graphicData")
    return first_element_with_tag(graphicdata, "chart")
end

"""
Concatenated text of a DrawingML rich-text block (`c:rich`, `cx:rich`), which is
`rich -> a:p* -> (a:r | a:fld)* -> a:t`.

Distinct from the shared-string/inline-string text of a worksheet cell, which
has no paragraph level. `nothing` rather than `""` when empty, so callers can
fall through to a cached title.
"""
function drawingml_text(rich::Union{Nothing,XML.Node})::Union{Nothing,String}
    isnothing(rich) && return nothing
    buf = IOBuffer()
    for p in XML.eachelement(rich)
        localname(p) == "p" || continue
        for run in XML.eachelement(p)
            localname(run) in ("r", "fld") || continue
            print(buf, something(child_text(run, "t"), ""))
        end
    end
    s = String(take!(buf))
    return isempty(s) ? nothing : s
end

"""
Resolve a `cx:f` source reference to a range.

Excel writes chartEx sources indirectly, through an auto-generated workbook
defined name (`_xlchart.v1.0`) rather than as a direct range, so a name is
resolved before parsing. Direct references are parsed as-is. `nothing` where
the name is absent or holds a constant rather than a reference.
"""
function resolve_chartex_ref(xf::XLSXFile, ref::AbstractString)::ChartRange
    r = parse_chart_range(ref)
    isnothing(r) || return r
    is_workbook_defined_name(xf, ref) || return nothing
    v = get_defined_name_value(get_workbook(xf), ref)
    return is_defined_name_value_a_reference(v) ? v : nothing
end

# ===========================================================================
# Cache parsing
# ===========================================================================

# A cached point is one of: a number, a known Excel error, or (rarely) text that
# is neither. Errors are recorded separately and stored as `missing`.
function parse_cached_point!(errors::Dict{Int,UInt64}, i::Int, s::AbstractString, numeric::Bool)
    if haskey(ERROR_STRING_TO_CODE, s)
        errors[i] = ERROR_STRING_TO_CODE[s]
        return missing
    end
    numeric || return String(s)
    v = tryparse(Float64, s)
    return isnothing(v) ? String(s) : v
end

"Dense vector of the `c:pt` children of a cache, literal or level node."
function parse_cached_points(node::XML.Node, numeric::Bool)
    n = child_val(node, "ptCount", 0)
    values = Vector{Any}(missing, n)
    errors = Dict{Int,UInt64}()
    for pt in XML.eachelement(node)
        localname(pt) == "pt" || continue
        i = tryparse(Int, get_attr(pt, "idx"))
        isnothing(i) && continue
        s = child_text(pt, "v")
        isnothing(s) && continue
        i += 1
        i < 1 && continue
        i > length(values) && append!(values, fill(missing, i - length(values)))  # tolerate a bad ptCount
        values[i] = parse_cached_point!(errors, i, s, numeric)
    end
    return (isempty(values) ? values : identity.(values)), errors
end

cache_ptcount(cache::Union{Nothing,XML.Node})::Int = child_val(cache, "ptCount", 0)

"""
Parse a series-role container (`c:tx`, `c:cat`, `c:val`, `c:xVal`, `c:yVal`,
`c:bubbleSize`) into a `ChartRef`, or `nothing` if it holds no cache.
"""
function parse_chart_ref(container::Union{Nothing,XML.Node}; read_cached_values::Bool=true)::Union{Nothing,ChartRef}
    isnothing(container) && return nothing
    for node in XML.eachelement(container)
        tag = localname(node)
        if tag == "numRef" || tag == "strRef"
            numeric = tag == "numRef"
            cnode = first_element_with_tag(node, numeric ? "numCache" : "strCache")
            pts, errs = (read_cached_values && !isnothing(cnode)) ? parse_cached_points(cnode, numeric) : (Any[], Dict{Int,UInt64}())
            return ChartRef(numeric ? :num : :str, child_text(node, "f"), child_text(cnode, "formatCode"),
                            cache_ptcount(cnode), pts, errs)
        elseif tag == "multiLvlStrRef"
            cnode = first_element_with_tag(node, "multiLvlStrCache")
            levels = Any[]
            if read_cached_values && !isnothing(cnode)
                for lvl in XML.eachelement(cnode)
                    localname(lvl) == "lvl" || continue
                    pts, _ = parse_cached_points(lvl, false)
                    push!(levels, pts)
                end
            end
            return ChartRef(:multiLvlStr, child_text(node, "f"), nothing, cache_ptcount(cnode), levels, Dict{Int,UInt64}())
        elseif tag == "numLit" || tag == "strLit"
            numeric = tag == "numLit"
            pts, errs = read_cached_values ? parse_cached_points(node, numeric) : (Any[], Dict{Int,UInt64}())
            return ChartRef(numeric ? :numLit : :strLit, nothing,
                            numeric ? child_text(node, "formatCode") : nothing,
                            cache_ptcount(node), pts, errs)
        elseif tag == "v"
            # <c:tx><c:v>Literal name</c:v></c:tx>
            return ChartRef(:strLit, nothing, nothing, 1,
                            Any[something(XML.is_simple_value(node), "")], Dict{Int,UInt64}())
        end
    end
    return nothing
end

function first_cached_string(r::Union{Nothing,ChartRef})::Union{Nothing,String}
    (isnothing(r) || isempty(r.data)) && return nothing
    v = first(r.data)
    (ismissing(v) || v isa AbstractVector) && return nothing
    return string(v)
end


# ===========================================================================
# External references
# ===========================================================================

# Replace the `[n]` index of an external workbook reference with the workbook
# path recorded in xl/externalLinks, as `getFormula(...; get_external_refs=true)`
# does for formulas.
function materialise_external_ref(xf::XLSXFile, ref::Union{Nothing,String})::Union{Nothing,String}
    isnothing(ref) && return nothing
    occursin('[', ref) || return ref
    out = ref
    for e in get_ext_refs(ref)
        path = try
            get_external_workbook_path(xf, e.index)
        catch err
            err isa XLSXError || rethrow()
            continue    # leave `[n]` unresolved rather than lose the chart
        end
        out = replace(out, "[" * string(e.index) * "]" => "[" * path * "]")
    end
    return out
end

function materialise(xf::XLSXFile, r::Union{Nothing,ChartRef})::Union{Nothing,ChartRef}
    isnothing(r) && return nothing
    isnothing(r.ref) && return r
    return ChartRef(r.kind, materialise_external_ref(xf, r.ref), r.format_code, r.ptCount, r.data, r.errors)
end


# ===========================================================================
# Part enumeration
# ===========================================================================

chart_name(path::AbstractString) = first(splitext(last(_split_zip_path(String(path)))))

# Part names declared in [Content_Types].xml with the given content type.
function parts_with_content_type(xf::XLSXFile, ctype::String)::Vector{String}
    paths = String[]
    haskey(xf.data, "[Content_Types].xml") || return paths
    for n in elements_with_tag(xml_root_element(xf.data["[Content_Types].xml"]), "Override")
        get_attr(n, "ContentType") == ctype || continue
        push!(paths, String(lstrip(get_attr(n, "PartName"), '/')))
    end
    return sort!(paths)
end

chart_parts(xf::XLSXFile)   = filter(p -> haskey(xf.data, p), parts_with_content_type(xf, CT_CHART))
chartex_parts(xf::XLSXFile) = filter(p -> haskey(xf.data, p), parts_with_content_type(xf, CT_CHARTEX))


# ===========================================================================
# Chart part parsing (c: schema)
# ===========================================================================

function parse_chart_series(ser::XML.Node, charttype::Symbol; read_cached_values::Bool=true)::ChartSeries
    # The series name always needs its cache, even under `read_cached_values=false`: it is metadata.
    name_ref = parse_chart_ref(first_element_with_tag(ser, "tx"); read_cached_values=true)

    categories = parse_chart_ref(first_element_with_tag(ser, "cat"); read_cached_values=read_cached_values)
    isnothing(categories) && (categories = parse_chart_ref(first_element_with_tag(ser, "xVal"); read_cached_values=read_cached_values))

    values = parse_chart_ref(first_element_with_tag(ser, "val"); read_cached_values=read_cached_values)
    isnothing(values) && (values = parse_chart_ref(first_element_with_tag(ser, "yVal"); read_cached_values=read_cached_values))

    return ChartSeries(
        child_val(ser, "idx", -1),
        child_val(ser, "order", -1),
        charttype,
        first_cached_string(name_ref),
        name_ref,
        categories,
        values,
        parse_chart_ref(first_element_with_tag(ser, "bubbleSize"); read_cached_values=read_cached_values),
    )
end

function parse_chart_title(chartnode::Union{Nothing,XML.Node})::Union{Nothing,String}
    tx = first_element_with_tag(first_element_with_tag(chartnode, "title"), "tx")
    isnothing(tx) && return nothing
    t = drawingml_text(first_element_with_tag(tx, "rich"))
    isnothing(t) || return t
    return first_cached_string(parse_chart_ref(tx; read_cached_values=true))
end

function parse_chart_part(
    xf::XLSXFile,
    path::String;
    read_cached_values::Bool=true,
    get_external_refs::Bool=false,
    rId::Union{Nothing,String}=nothing,
    sheet::Union{Nothing,String}=nothing,
    from::Union{Nothing,String}=nothing,
    to::Union{Nothing,String}=nothing,
)::Chart

    haskey(xf.data, path) || throw(XLSXError("Chart part `$path` not found in the package."))

    chartspace = xml_root_element(xf.data[path])
    localname(chartspace) != "chartSpace" &&
        throw(XLSXError("Malformed chart part $path. Root node name should be `chartSpace`. Found $(localname(chartspace))."))

    chartnode = first_element_with_tag(chartspace, "chart")
    plotarea = first_element_with_tag(chartnode, "plotArea")

    charttypes = Symbol[]
    series = ChartSeries[]
    for group in (isnothing(plotarea) ? () : XML.eachelement(plotarea))
        tag = localname(group)
        tag in CHART_GROUP_TAGS || continue
        charttype = Symbol(tag)
        push!(charttypes, charttype)
        for ser in elements_with_tag(group, "ser")
            s = parse_chart_series(ser, charttype; read_cached_values=read_cached_values)
            if get_external_refs
                s = ChartSeries(s.idx, s.order, s.charttype, s.name,
                                materialise(xf, s.name_ref),
                                materialise(xf, s.categories),
                                materialise(xf, s.values),
                                materialise(xf, s.bubble_sizes))
            end
            push!(series, s)
        end
    end

    _, fname = _split_zip_path(path)
    return Chart(path, first(splitext(fname)), rId, sheet, from, to,
                 parse_chart_title(chartnode), charttypes, series)
end


# ===========================================================================
# Chart part parsing (cx: schema)
# ===========================================================================

function parse_chartex_part(
    xf::XLSXFile,
    path::String;
    rId::Union{Nothing,String}=nothing,
    sheet::Union{Nothing,String}=nothing,
    from::Union{Nothing,String}=nothing,
    to::Union{Nothing,String}=nothing,
)::ChartEx

    haskey(xf.data, path) || throw(XLSXError("Chart part `$path` not found in the package."))

    chartspace = xml_root_element(xf.data[path])
    localname(chartspace) != "chartSpace" &&
        throw(XLSXError("Malformed chartEx part $path. Root node name should be `chartSpace`. Found $(localname(chartspace))."))

    # Source ranges live up front in cx:chartData, shared across series, rather
    # than being carried by each series as in the c: schema.
    refs = String[]
    ranges = ChartRange[]
    chartdata = first_element_with_tag(chartspace, "chartData")
    if !isnothing(chartdata)
        for d in elements_with_tag(chartdata, "data")
            for dim in XML.eachelement(d)
                localname(dim) in ("numDim", "strDim") || continue
                fml = child_text(dim, "f")
                isnothing(fml) && continue
                push!(refs, fml)
                push!(ranges, resolve_chartex_ref(xf, fml))
            end
        end
    end

    chartnode = first_element_with_tag(chartspace, "chart")
    plotarea = first_element_with_tag(chartnode, "plotArea")
    region = first_element_with_tag(plotarea, "plotAreaRegion")

    # The chart type is an attribute here, not an element name.
    layouts = String[]
    binning = false
    for ser in (isnothing(region) ? XML.Node[] : elements_with_tag(region, "series"))
        lid = get_attr(ser, "layoutId")
        isempty(lid) || push!(layouts, lid)
        isnothing(first_element_with_tag(ser, "binning")) || (binning = true)
    end

    title = drawingml_text(first_element_with_tag(
        first_element_with_tag(first_element_with_tag(chartnode, "title"), "tx"), "rich"))

    _, fname = _split_zip_path(path)
    return ChartEx(path, first(splitext(fname)), rId, sheet, from, to,
                   title, layouts, refs, ranges, binning)
end


# ===========================================================================
# Discovery: sheet -> drawing -> chart
# ===========================================================================

function charts_for_sheet!(found::Vector{ChartLocation}, xf::XLSXFile,
                           sheet_path::String, sheet_name::String)
    drawing_path = _drawing_path_for_sheet(xf, sheet_path)
    isnothing(drawing_path) && return found
    haskey(xf.data, drawing_path) || return found

    rid_map = rid_to_target(xf, drawing_path, (REL_CHART, REL_CHARTEX))
    isempty(rid_map) && return found

    for anchor in XML.eachelement(xml_root_element(xf.data[drawing_path]))
        endswith(localname(anchor), "Anchor") || continue
        from = _parse_cell_marker(anchor, "from"; is_to=false)
        to   = _parse_cell_marker(anchor, "to"; is_to=true)
        for shape in effective_shapes(anchor)
            chart_el = frame_chart_element(shape)
            isnothing(chart_el) && continue
            rId = get_prefixed_attr(chart_el, "id")
            isnothing(rId) && continue
            entry = get(rid_map, rId, nothing)
            isnothing(entry) && continue        # a picture or shape, not a chart
            chart_path, reltype = entry
            haskey(xf.data, chart_path) || continue
            push!(found, ChartLocation(chart_path,
                                       (sheet=sheet_name, from=from, to=to, rId=rId),
                                       reltype == REL_CHARTEX ? :cx : :c))
        end
    end

    return found
end

function chart_anchors(xf::XLSXFile)::Vector{ChartLocation}
    wb = get_workbook(xf)
    found = ChartLocation[]
    for sheet in wb.sheets
        sheet_path = get_relationship_target_by_id("xl", wb, sheet.relationship_id)
        charts_for_sheet!(found, xf, sheet_path, sheet.name)
    end
    return found
end

# Chart parts in document order, each with its anchor when a drawing references
# it. Sheet-anchored charts come first (sheet order, then anchor order); parts
# the package declares but no drawing references follow.
function chart_positions(xf::XLSXFile)::Vector{ChartLocation}
    out = ChartLocation[]
    seen = Set{String}()
    for loc in chart_anchors(xf)
        loc.path in seen && continue
        push!(seen, loc.path)
        push!(out, loc)
    end
    for (path, sch) in Iterators.flatten((
            ((p, :c)  for p in chart_parts(xf)),
            ((p, :cx) for p in chartex_parts(xf)),
        ))
        path in seen && continue
        push!(seen, path)
        push!(out, ChartLocation(path, nothing, sch))
    end
    return out
end

function chart_positions(ws::Worksheet)::Vector{ChartLocation}
    xf = get_xlsxfile(ws)
    found = ChartLocation[]
    sheet_path = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    charts_for_sheet!(found, xf, sheet_path, ws.name)
    out = ChartLocation[]
    seen = Set{String}()
    for loc in found
        loc.path in seen && continue
        push!(seen, loc.path)
        push!(out, loc)
    end
    return out
end

function parse_chart_at(xf::XLSXFile, loc::ChartLocation; kw...)::AbstractChart
    a = loc.anchor
    if loc.schema === :cx
        return isnothing(a) ? parse_chartex_part(xf, loc.path) :
                              parse_chartex_part(xf, loc.path; a...)
    end
    return isnothing(a) ? parse_chart_part(xf, loc.path; kw...) :
                          parse_chart_part(xf, loc.path; kw..., a...)
end

# Excel names a chartEx chart's sources `_xlchart.v<N>.<M>`, where N is a
# counter it bumps per batch of charts and M runs within the batch. Pick a
# fresh N so a copied chart's names cannot collide with the original's.
function _next_xlchart_series(wb::Workbook)::Int
    n = 0
    for name in keys(wb.workbook_names)
        m = match(r"^_xlchart\.v(\d+)\.\d+$", name)
        isnothing(m) && continue
        n = max(n, parse(Int, m.captures[1]))
    end
    return n + 1
end

# ===========================================================================
# Accessors common to both schemas
# ===========================================================================

"""
    chartSchema(c::AbstractChart) -> Symbol

`:c` for charts in the original schema, `:cx` for newer `chartEx` charts.
"""
chartSchema(::Chart)   = :c
chartSchema(::ChartEx) = :cx

"""
    chartType(c::AbstractChart) -> Symbol

The kind of chart, as a single symbol, whatever the schema.

For `c:` charts this is the plot-group tag - `:barChart`, `:scatterChart` - or
`:combo` where the chart has more than one group. The full list is in
`c.charttypes`, and each series carries its own group in `s.charttype`.

For `cx:` charts it is the normalised layout: `:waterfall`, `:funnel`,
`:treemap`, `:sunburst`, `:boxWhisker`, `:regionMap`, `:histogram`, `:pareto`,
or `:clusteredColumn`. Histogram and Pareto are not distinct layouts in the
file - a histogram is a clustered column series carrying `cx:binning`, and a
Pareto adds a second `paretoLine` series - so both are derived here.
"""
function chartType(c::Chart)::Symbol
    isempty(c.charttypes) && return :unknown
    return length(c.charttypes) == 1 ? only(c.charttypes) : :combo
end

function chartType(c::ChartEx)::Symbol
    ls = c.layouts
    isempty(ls) && return :unknown
    "paretoLine" in ls && return :pareto
    length(ls) == 1 || return :combo
    only(ls) == "clusteredColumn" && return c.binning ? :histogram : :clusteredColumn
    return Symbol(only(ls))
end

chartpath(c::AbstractChart)  = c.path
chartname(c::AbstractChart)  = c.name
charttitle(c::AbstractChart) = c.title
sheetname(c::AbstractChart)  = c.sheet


# ===========================================================================
# Public API
# ===========================================================================

"""
    getCharts(xf::XLSXFile; read_cached_values=true, get_external_refs=false) -> Vector{AbstractChart}
    getCharts(ws::Worksheet; read_cached_values=true, get_external_refs=false) -> Vector{AbstractChart}

Return every chart in the file, or every chart anchored to `ws`, together with
the data Excel cached inside each chart part.

Pass `read_cached_values=false` to read metadata only - title, chart types, series names,
source formulas, format codes and point counts - and skip the cached values,
which is the expensive part for a large chart.

A chart may reference an external workbook, in which case its source formula
takes the form `[1]Sheet1!\$A\$1:\$A\$10`, where `[1]` indexes the workbook's
external references. Use `get_external_refs=true` to substitute the workbook
path, as [`XLSX.getFormula`](@ref) does.

# Examples
```julia
julia> f = XLSX.readxlsx("sales.xlsx");

julia> c = XLSX.getCharts(f["Summary"])[1];

julia> c.title
"Revenue by region"

julia> c.series[1].values.ref
"Summary!\$B\$2:\$B\$5"

# To create a set of DataFrames from chart data skiping ChartEx
julia> dfs = Dict(c.name => DataFrame(XLSX.getChartData(c))
                  for c in XLSX.getCharts(f) if c isa XLSX.Chart)
```

!!! note
    The values returned are Excel's cache, written when the file was last saved
    by Excel. They may be stale relative to the source, and a file written by a
    tool that does not populate the cache will return empty series.

!!! note
    Charts using the newer `chartEx` schema - waterfall, funnel, treemap,
    sunburst, histogram, Pareto, box & whisker, region map - are returned as
    [`XLSX.ChartEx`](@ref) rather than [`XLSX.Chart`](@ref). Their type, title
    and source ranges are available; their cached values are not, and
    [`XLSX.getChartData`](@ref) throws for them. Use [`XLSX.chartSchema`](@ref) or
    `isa` to tell the two apart. `read_cached_values` and `get_external_refs`
    have no effect on them.

See also [`XLSX.getChart`](@ref), [`XLSX.getChartData`](@ref), [`XLSX.chartType`](@ref).
"""
function getCharts(x::Union{Worksheet,XLSXFile};
                   read_cached_values::Bool=true,
                   get_external_refs::Bool=false)::Vector{AbstractChart}
    xf = get_xlsxfile(x)
    charts = AbstractChart[]
    for loc in chart_positions(x)
        push!(charts, parse_chart_at(xf, loc;
                                     read_cached_values=read_cached_values,
                                     get_external_refs=get_external_refs))
    end
    return charts
end

"""
    getChart(ws::Worksheet, name; read_cached_values=true, get_external_refs=false) -> AbstractChart
    getChart(xf::XLSXFile, name; read_cached_values=true, get_external_refs=false) -> AbstractChart

Return a single chart. `name` may be the part name (`"chart1"` or
`"chart1.xml"`), the full package path, or the chart's relationship id within its
drawing part (`"rId1"`).

Returns a [`XLSX.ChartEx`](@ref) where the named part uses the `chartEx` schema.

See also [`XLSX.getCharts`](@ref).
"""
function getChart(x::Union{Worksheet,XLSXFile}, name::AbstractString;
                  read_cached_values::Bool=true,
                  get_external_refs::Bool=false)::AbstractChart
    xf = get_xlsxfile(x)
    positions = chart_positions(x)
    stem = chart_name(name)
    for loc in positions
        loc.path == name || chart_name(loc.path) == stem ||
            (!isnothing(loc.anchor) && loc.anchor.rId == name) || continue
        return parse_chart_at(xf, loc;
                              read_cached_values=read_cached_values,
                              get_external_refs=get_external_refs)
    end
    throw(XLSXError("No chart matching `$name`. Found: " *
                    join((chart_name(l.path) for l in positions), ", ") * "."))
end


# ===========================================================================
# Errors in cached values
# ===========================================================================

"""
    iserror(r::ChartRef) -> Vector{Bool}
    iserror(r::ChartRef, i::Integer) -> Bool

Report which cached chart values are Excel error values. When Excel writes source data 
to a chart cache, all error values are written as simple zeros except for `#N/A`. 
Therefore, when  operating on a chart cache, only `#N/A` values will return `true`. 
All other error values will return `false`, and are indistinguisable from genuine 
zero values.

See also [`XLSX.geterror`](@ref).
"""
iserror(r::ChartRef)::Vector{Bool} = Bool[haskey(r.errors, i) for i in 1:length(r.data)]
iserror(r::ChartRef, i::Integer)::Bool = haskey(r.errors, Int(i))

"""
    geterror(r::ChartRef) -> Vector{String}
    geterror(r::ChartRef, i::Integer) -> String

Resolve cached chart `#N/A`error values to their Excel strings (`"#N/A"`). All other 
error values are written by Excel as simple zeros in the chart cache, and return "".

See also [`XLSX.iserror`](@ref).
"""
geterror(r::ChartRef)::Vector{String} = String[geterror(r, i) for i in 1:length(r.data)]
geterror(r::ChartRef, i::Integer)::String =
    haskey(r.errors, Int(i)) ? get_error_string(r.errors[Int(i)]) : ""


# ===========================================================================
# Cached data as a table
# ===========================================================================

function unique_label!(labels::Vector{Symbol}, name::AbstractString)::Symbol
    base = Symbol(isempty(name) ? "column" : name)
    label = base
    n = 1
    while label in labels
        n += 1
        label = Symbol(base, "_", n)
    end
    push!(labels, label)
    return label
end

pad_to(v::AbstractVector, n::Int) =
    length(v) >= n ? collect(v) : vcat(collect(v), fill(missing, n - length(v)))

# A ref read with `cache=false` keeps its declared ptCount but no values. An
# absent cache gives ptCount == 0; a cache of blanks gives a full-length vector
# of `missing`. So this combination is unambiguous.
no_cached_values(r::ChartRef)::Bool = r.ptCount > 0 && isempty(r.data)

"""
    getChartData(c::Chart) -> DataTable
    getChartData(ws::Worksheet, name) -> DataTable
    getChartData(xf::XLSXFile, name) -> DataTable

Return the cached data of a chart as a `DataTable`, ready for `DataFrame(...)` or
any other Tables.jl sink.

Categories become the leading column(s). Where every series shares one category
reference a single `categories` column is emitted; otherwise each series
contributes its own `<series>_x` column, which is the usual layout for scatter
and bubble charts. Multi-level categories give one column per level, and bubble
charts add a `<series>_size` column. Series of unequal length are padded with
`missing`.

Series with no name in the file - Excel shows these as "Series1", "Series2" in
the legend - are labelled by position. No category column is produced when the
chart has no `c:cat` at all: Excel is plotting against an implicit index in that
case, and nothing is cached for it.

# Examples
```julia
julia> using DataFrames

julia> DataFrame(XLSX.getChartData(f["Summary"], "chart1"))
4×3 DataFrame
 Row │ categories  2024      2025
     │ String      Float64   Float64
```

!!! note
    `chartEx` charts carry no readable value cache, so this throws
    [`XLSX.XLSXError`](@ref) for them. Use
    [`XLSX.getChartRanges`](@ref) and [`XLSX.getdata`](@ref) to read their
    source cells instead.

See also [`XLSX.getCharts`](@ref), [`XLSX.gettable`](@ref).
"""
function getChartData(c::Chart)::DataTable
    isempty(c.series) && return DataTable(Any[], Symbol[])

    for s in c.series, r in (s.categories, s.values, s.bubble_sizes)
        isnothing(r) && continue
        no_cached_values(r) && throw(XLSXError(
            "Chart `$(c.name)` was read with `read_cached_values=false`, so its cached values are not available. Read it again with `read_cached_values=true`."))
    end
    catrefs = [s.categories for s in c.series]
    shared = !isnothing(first(catrefs)) && !isnothing(first(catrefs).ref) &&
             all(r -> !isnothing(r) && r.ref == first(catrefs).ref, catrefs)

    n = 0
    for s in c.series, r in (s.categories, s.values)
        isnothing(r) && continue
        n = max(n, r.ptCount, length(r.data))
    end

    columns = Any[]
    labels = Symbol[]

    function push_categories!(r::ChartRef, prefix::AbstractString)
        if r.kind == :multiLvlStr
            for (i, level) in enumerate(r.data)
                unique_label!(labels, "$(prefix)_$(i)")
                push!(columns, pad_to(level, n))
            end
        else
            unique_label!(labels, prefix)
            push!(columns, pad_to(r.data, n))
        end
    end

    shared && push_categories!(first(catrefs), "categories")

    for (i, s) in enumerate(c.series)
        name = something(s.name, "Series$(i)")
        shared || isnothing(s.categories) || push_categories!(s.categories, "$(name)_x")
        unique_label!(labels, name)
        push!(columns, isnothing(s.values) ? Vector{Any}(missing, n) : pad_to(s.values.data, n))
        if !isnothing(s.bubble_sizes)
            unique_label!(labels, "$(name)_size")
            push!(columns, pad_to(s.bubble_sizes.data, n))
        end
    end

    return DataTable(columns, labels)
end

"""
    getChartData(c::ChartEx)

`chartEx` charts do not carry a readable value cache. Throws
[`XLSX.XLSXError`](@ref).
"""
getChartData(c::ChartEx) =
    throw(XLSXError("Cannot get data for chart `$(c.name)`: it is a $(chartType(c)) chart, " *
                    "which uses the `chartEx` schema and carries no readable value cache. " *
                    "Use `getChartRanges` to find its source cells and `getdata` to read them."))

getChartData(x::Union{Worksheet,XLSXFile}, name::AbstractString; kw...)::DataTable =
    getChartData(getChart(x, name; kw...))


# ===========================================================================
# Source ranges
# ===========================================================================

"""
    parse_chart_range(ref) -> ChartRange

A chart source formula as a range, or `nothing` when it has no addressable one:
literal series, external-workbook references, and defined names.
"""
function parse_chart_range(ref::Union{Nothing,AbstractString})
    isnothing(ref) && return nothing
    s = strip(ref)
    isempty(s) && return nothing
    occursin('[', s) && return nothing                  # external workbook
    if startswith(s, '(') && endswith(s, ')')           # multi-area
        s = s[nextind(s, firstindex(s)):prevind(s, lastindex(s))]
    end
    occursin(',', s) && return NonContiguousRange(String(s))
    (is_valid_fixed_sheet_cellrange(s)    || is_valid_sheet_cellrange(s))    && return SheetCellRange(s)
    (is_valid_fixed_sheet_cellname(s)     || is_valid_sheet_cellname(s))     && return SheetCellRef(s)
    (is_valid_fixed_sheet_column_range(s) || is_valid_sheet_column_range(s)) && return SheetColumnRange(s)
    (is_valid_fixed_sheet_row_range(s)    || is_valid_sheet_row_range(s))    && return SheetRowRange(s)
    return nothing                                      # defined name, or unrecognised
end

"""
    chart_range(r::ChartRef) -> ChartRange

The source range of a `ChartRef`, or `nothing` when it has none.
"""
chart_range(r::Union{Nothing,ChartRef}) =
    isnothing(r) ? nothing : parse_chart_range(r.ref)

"""
    getChartRanges(c::Chart) -> Vector{ChartRanges}
    getChartRanges(c::ChartEx) -> Vector{ChartRange}
    getChartRanges(ws::Worksheet, name) -> Vector
    getChartRanges(xf::XLSXFile, name) -> Vector
    getChartRanges(ws::Worksheet) -> Vector{@NamedTuple{chart::String, ranges::Vector}}
    getChartRanges(xf::XLSXFile) -> Vector{@NamedTuple{chart::String, ranges::Vector}}

The worksheet ranges of the source data a chart plots from.

Given a [`XLSX.Chart`](@ref), or a chart `name` in any of the forms
[`XLSX.getChart`](@ref) accepts, return one entry per series in document order,
parallel to `c.series`. Each entry carries the series `idx` and `name` alongside
its `categories`, `values` and `bubble_sizes` ranges.

Given no name, return the ranges of every chart on the worksheet or in the
workbook, each paired with its chart name, following [`XLSX.getCharts`](@ref).

`categories` holds `c:cat` or `c:xVal` and `values` holds `c:val` or `c:yVal`, so
the two mean the same thing whatever the chart type, as in [`XLSX.ChartSeries`](@ref).
`bubble_sizes` is `nothing` for every chart type but bubble.

A range is `nothing` wherever the series has no addressable source: a literal
series (`c:numLit`/`c:strLit`), a reference to an external workbook, a defined
name.

# Examples
```julia
julia> f = XLSX.readxlsx("sales.xlsx");

julia> r = XLSX.getChartRanges(f["Summary"], "chart1");

julia> r[1].name, r[1].values
("2024", Summary!B2:B5)

julia> XLSX.getdata(f, r[1].values)      # read the live source cells, not the cache
4-element Vector{Any}:
 1250.0
 1310.0
 ⋮

julia> [(x.chart, length(x.ranges)) for x in XLSX.getChartRanges(f)]
2-element Vector{Tuple{String, Int64}}:
 ("chart1", 3)
 ("chart2", 1)
```

!!! note
    A range records where the chart says its source data came from, which is not
    necessarily where the values in [`XLSX.getChartData`](@ref) came from: the
    cache is a snapshot from the last save, and the cells may have changed
    since, or the source sheet may have been deleted entirely.

!!! note
    For a [`XLSX.ChartEx`](@ref) this returns a flat `Vector{ChartRange}` rather
    than one entry per series. The `cx:` schema declares its data dimensions
    once in `cx:chartData` and shares them across series, so there is no
    per-series `idx` or `name` to report.

See also [`XLSX.getChart`](@ref), [`XLSX.getCharts`](@ref), [`XLSX.getChartData`](@ref).
"""
getChartRanges(c::Chart)::Vector{ChartRanges} =
    [(idx = s.idx,
      name = s.name,
      categories = chart_range(s.categories),
      values = chart_range(s.values),
      bubble_sizes = chart_range(s.bubble_sizes))
     for s in c.series]

getChartRanges(c::ChartEx)::Vector{ChartRange} = c.ranges

getChartRanges(x::Union{Worksheet,XLSXFile}, name::AbstractString) =
    getChartRanges(getChart(x, name; read_cached_values=false))

getChartRanges(x::Union{Worksheet,XLSXFile}) =
    [(chart = c.name, ranges = getChartRanges(c)) for c in getCharts(x; read_cached_values=false)]


# ===========================================================================
# Display
# ===========================================================================

function Base.show(io::IO, c::Chart)
    print(io, "XLSX.Chart(\"", c.name, "\"",
          isnothing(c.sheet) ? "" : ", \"" * c.sheet * "\"",
          isnothing(c.from) ? "" : ", " * c.from,
          ", ", join(string.(c.charttypes), "+"),
          ", ", length(c.series), " series)")
end

Base.show(io::IO, s::ChartSeries) =
    print(io, "XLSX.ChartSeries(", something(s.name, "<unnamed>"), ", ", s.charttype, ")")

Base.show(io::IO, r::ChartRef) =
    print(io, "XLSX.ChartRef(", something(r.ref, "<literal>"), ", ", r.ptCount, " pts)")

function Base.show(io::IO, ::MIME"text/plain", c::Chart)
    print(io, "XLSX.Chart \"", c.name, "\"")
    isnothing(c.sheet) || print(io, " on sheet \"", c.sheet, "\"")
    isnothing(c.from) || print(io, " at ", c.from, isnothing(c.to) ? "" : ":" * c.to)
    println(io)
    isnothing(c.title) || println(io, "  title: ", repr(c.title))
    println(io, "  type: ", join(string.(c.charttypes), ", "))
    println(io, "  series: ", length(c.series))
    for (i, s) in enumerate(c.series)
        print(io, "    [", s.order, "] ", something(s.name, "Series$(i)"))
        isnothing(s.values) ||
            print(io, " - ", something(s.values.ref, "<literal>"), " (", s.values.ptCount, " pts)")
        println(io)
    end
end

function Base.show(io::IO, ::MIME"text/plain", s::ChartSeries)
    println(io, "XLSX.ChartSeries ", something(s.name, "<unnamed>"), " (", s.charttype, ")")
    for (label, r) in (("categories", s.categories), ("values", s.values), ("sizes", s.bubble_sizes))
        isnothing(r) && continue
        println(io, "  ", label, ": ", something(r.ref, "<literal>"), " - ", r.ptCount, " pts, ", r.kind)
    end
end

function Base.show(io::IO, ::MIME"text/plain", r::ChartRef)
    println(io, "XLSX.ChartRef ", something(r.ref, "<literal>"), " (", r.kind, ", ", r.ptCount, " pts)")
    isnothing(r.format_code) || println(io, "  format: ", r.format_code)
    isempty(r.errors) || println(io, "  errors at: ", join(sort(collect(keys(r.errors))), ", "))
    isempty(r.data) || println(io, "  data: ", r.data)
end

Base.show(io::IO, c::ChartEx) =
    print(io, "XLSX.ChartEx(\"", c.name, "\"",
          isnothing(c.sheet) ? "" : ", \"" * c.sheet * "\"",
          isnothing(c.from) ? "" : ", " * c.from,
          ", ", chartType(c), ", ", length(c.refs), " refs)")

function Base.show(io::IO, ::MIME"text/plain", c::ChartEx)
    print(io, "XLSX.ChartEx \"", c.name, "\"")
    isnothing(c.sheet) || print(io, " on sheet \"", c.sheet, "\"")
    isnothing(c.from) || print(io, " at ", c.from, isnothing(c.to) ? "" : ":" * c.to)
    println(io)
    isnothing(c.title) || println(io, "  title: ", repr(c.title))
    println(io, "  type: ", chartType(c))
    isempty(c.layouts) || println(io, "  layouts: ", join(c.layouts, ", "))
    println(io, "  refs: ", isempty(c.refs) ? "none" : join(c.refs, ", "))
    println(io, "  (chartEx: cached values and appearance are not read)")
end
