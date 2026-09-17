# Child order for xsd:sequence elements in the c: chart schema (ECMA-376,
# dml-chart.xsd), denormalized — EG_SerShared and EG_AxShared expanded inline.
# An element inserted out of this order makes the part invalid and Excel will
# refuse or repair the file.
#
# mc:AlternateContent may appear anywhere and has no slot, so _schema_position
# skips it. Excel uses it in c:chartSpace to wrap c:style. A setter for an
# element Excel wraps this way must look inside the AlternateContent rather than
# insert a sibling, or the file ends up with two competing values.
#
# XML.jl's Node is a non-mutable struct whose `children` and `attributes` fields
# are `Union{Nothing,Vector}`. push! and setindex! exist and mutate in place, but
# only when the field is already a vector — on an element parsed from `<a/>` they
# throw "Node does not accept children", because the struct cannot be given one.
# Excel writes empty elements routinely, so every write here returns a new node
# rather than mutating: insert_child, replace_child, remove_child and
# with_attribute all take a node and give one back.
#
const CHILD_ORDER = Dict{Tuple{String,String},Vector{Union{String,Vector{String}}}}(

    # Series, one per chart type
    (NS_C, "areaSer")    => ["idx","order","tx","spPr","pictureOptions","dPt","dLbls",
                             "trendline","errBars","cat","val","extLst"],
    (NS_C, "barSer")     => ["idx","order","tx","spPr","invertIfNegative","pictureOptions",
                             "dPt","dLbls","trendline","errBars","cat","val","shape","extLst"],
    (NS_C, "bubbleSer")  => ["idx","order","tx","spPr","invertIfNegative","dPt","dLbls",
                             "trendline","errBars","xVal","yVal","bubbleSize","bubble3D","extLst"],
    (NS_C, "lineSer")    => ["idx","order","tx","spPr","marker","dPt","dLbls","trendline",
                             "errBars","cat","val","smooth","extLst"],
    (NS_C, "pieSer")     => ["idx","order","tx","spPr","explosion","dPt","dLbls",
                             "cat","val","extLst"],
    (NS_C, "radarSer")   => ["idx","order","tx","spPr","marker","dPt","dLbls",
                             "cat","val","extLst"],
    (NS_C, "scatterSer") => ["idx","order","tx","spPr","marker","dPt","dLbls","trendline",
                             "errBars","xVal","yVal","smooth","extLst"],
    (NS_C, "surfaceSer") => ["idx","order","tx","spPr","cat","val","extLst"],

    # Points and markers
    (NS_C, "dPt")    => ["idx","invertIfNegative","marker","bubble3D","explosion","spPr",
                         "pictureOptions","extLst"],
    (NS_C, "marker") => ["symbol","size","spPr","extLst"],

    # Data labels
    (NS_C, "dLbls") => ["dLbl","delete","numFmt","spPr","txPr","dLblPos","showLegendKey",
                        "showVal","showCatName","showSerName","showPercent","showBubbleSize",
                        "separator","showLeaderLines","leaderLines","extLst"],
    (NS_C, "dLbl")  => ["idx","delete","layout","tx","numFmt","spPr","txPr","dLblPos",
                        "showLegendKey","showVal","showCatName","showSerName","showPercent",
                        "showBubbleSize","separator","extLst"],

    # Axes
    (NS_C, "catAx") => ["axId","scaling","delete","axPos","majorGridlines","minorGridlines",
                        "title","numFmt","majorTickMark","minorTickMark","tickLblPos","spPr",
                        "txPr","crossAx","crosses","crossesAt","auto","lblAlgn","lblOffset",
                        "tickLblSkip","tickMarkSkip","noMultiLvlLbl","extLst"],
    (NS_C, "valAx") => ["axId","scaling","delete","axPos","majorGridlines","minorGridlines",
                        "title","numFmt","majorTickMark","minorTickMark","tickLblPos","spPr",
                        "txPr","crossAx","crosses","crossesAt","crossBetween","majorUnit",
                        "minorUnit","dispUnits","extLst"],
    (NS_C, "dateAx") => ["axId","scaling","delete","axPos","majorGridlines","minorGridlines",
                         "title","numFmt","majorTickMark","minorTickMark","tickLblPos","spPr",
                         "txPr","crossAx","crosses","crossesAt","auto","lblOffset","baseTimeUnit",
                         "majorUnit","majorTimeUnit","minorUnit","minorTimeUnit","extLst"],
    (NS_C, "serAx") => ["axId","scaling","delete","axPos","majorGridlines","minorGridlines",
                        "title","numFmt","majorTickMark","minorTickMark","tickLblPos","spPr",
                        "txPr","crossAx","crosses","crossesAt","tickLblSkip","tickMarkSkip","extLst"],

    # CT_ChartLines — gridlines, drop lines, hi-low lines, series lines, leader lines
    (NS_C, "majorGridlines") => ["spPr"],
    (NS_C, "minorGridlines") => ["spPr"],
    (NS_C, "dropLines")      => ["spPr"],
    (NS_C, "hiLowLines")     => ["spPr"],
    (NS_C, "serLines")       => ["spPr"],
    (NS_C, "leaderLines")    => ["spPr"],

    # Shape and text properties. Keyed on NS_A: c:spPr is a:CT_ShapeProperties and
    # c:txPr / c:rich are a:CT_TextBody — DrawingML types used at chart-prefixed
    # elements — so the child order belongs to DrawingML regardless of the prefix
    # the element itself carries.
    (NS_A, "spPr") => ["xfrm","custGeom","prstGeom","noFill","solidFill","gradFill",
                       "blipFill","pattFill","grpFill","ln","effectLst","effectDag",
                       "scene3d","sp3d","extLst"],
    (NS_A, "txPr") => ["bodyPr","lstStyle","p"],
    (NS_A, "rich") => ["bodyPr","lstStyle","p"],

    # CT_Tx is a chart type, unlike the two above.
    (NS_C, "tx")   => ["strRef","rich"],

    # Chart space and plot area
    (NS_C, "chartSpace") => ["date1904","lang","roundedCorners","style","clrMapOvr",
                             "pivotSource","protection","chart","spPr","txPr",
                             "externalData","printSettings","userShapes","extLst"],
    (NS_C, "chart")      => ["title","autoTitleDeleted","pivotFmts","view3D","floor",
                             "sideWall","backWall","plotArea","legend","plotVisOnly",
                             "dispBlanksAs","showDLblsOverMax","extLst"],
    (NS_C, "plotArea") => ["layout",
                           ["areaChart","area3DChart","lineChart","line3DChart","stockChart",
                            "radarChart","scatterChart","pieChart","pie3DChart","doughnutChart",
                            "barChart","bar3DChart","ofPieChart","surfaceChart","surface3DChart",
                            "bubbleChart"],
                           ["valAx","catAx","dateAx","serAx"],
                           "dTable","spPr","extLst"],

    # Groups
    (NS_C, "barChart")  => ["barDir","grouping","varyColors","ser","dLbls","gapWidth",
                            "overlap","serLines","axId","extLst"],
    (NS_C, "lineChart") => ["grouping","varyColors","ser","dLbls","dropLines","hiLowLines",
                            "upDownBars","marker","smooth","axId","extLst"],

    # CT_LineProperties. Three choice groups then the ends. Note the line fill
    # group has four members, not the six of CT_ShapeProperties — a line cannot
    # take a blipFill or grpFill.
    (NS_A, "ln") => ["noFill","solidFill","gradFill","pattFill",
                     "prstDash","custDash",
                     "round","bevel","miter",
                     "headEnd","tailEnd","extLst"],

    # Text body internals. CT_TextCharacterProperties serves a:defRPr, a:rPr and
    # a:endParaRPr — the run properties themselves are attributes on it, so this
    # order governs only its fill, line and typeface children.
    (NS_A, "p")    => ["pPr", "r", "br", "fld", "endParaRPr"],
    (NS_A, "pPr")  => ["lnSpc","spcBef","spcAft","buClrTx","buClr","buSzTx","buSzPct",
                       "buSzPts","buFontTx","buFont","buNone","buAutoNum","buChar",
                       "buBlip","tabLst","defRPr","extLst"],
    (NS_A, "defRPr") => ["ln","noFill","solidFill","gradFill","blipFill","pattFill",
                         "grpFill","effectLst","effectDag","highlight","uLnTx","uLn",
                         "uFillTx","uFill","latin","ea","cs","sym","hlinkClick",
                         "hlinkMouseOver","rtl","extLst"],
    (NS_A, "r")   => ["rPr", "t"],
    (NS_A, "fld") => ["rPr", "pPr", "t"],
    (NS_A, "br")  => ["rPr"],
    (NS_A, "lnSpc")  => [["spcPct", "spcPts"]],
    (NS_A, "spcBef") => [["spcPct", "spcPts"]],
    (NS_A, "spcAft") => [["spcPct", "spcPts"]],
    (NS_A, "bodyPr") => ["prstTxWarp",
                         "noAutofit", "normAutofit", "spAutoFit",
                         "scene3d",
                         "sp3d", "flatTx",
                         "extLst"],
    (NS_C, "title")       => ["tx","layout","overlay","spPr","txPr","extLst"],
    (NS_C, "legend")      => ["legendPos","legendEntry","layout","overlay","spPr",
                              "txPr","extLst"],
    (NS_C, "legendEntry") => ["idx","delete","txPr","extLst"],          
)

CHILD_ORDER[(NS_A, "rPr")]         = CHILD_ORDER[(NS_A, "defRPr")]
CHILD_ORDER[(NS_A, "endParaRPr")]  = CHILD_ORDER[(NS_A, "defRPr")]

# Tags that may appear more than once, keyed by parent. errBars is maxOccurs=2
# on area, bubble and scatter series and 1 elsewhere, so this cannot be a flat set.
const REPEATABLE = Dict{Tuple{String,String},Set{String}}(
    (NS_C, "areaSer")    => Set(["dPt", "trendline", "errBars"]),
    (NS_C, "barSer")     => Set(["dPt", "trendline"]),
    (NS_C, "bubbleSer")  => Set(["dPt", "trendline", "errBars"]),
    (NS_C, "lineSer")    => Set(["dPt", "trendline"]),
    (NS_C, "pieSer")     => Set(["dPt"]),
    (NS_C, "radarSer")   => Set(["dPt"]),
    (NS_C, "scatterSer") => Set(["dPt", "trendline", "errBars"]),
    (NS_C, "surfaceSer") => Set{String}(),
    (NS_C, "dLbls")      => Set(["dLbl"]),
    (NS_A, "txPr")       => Set(["p"]),
    (NS_A, "rich")       => Set(["p"]),
    (NS_C, "barChart")  => Set(["ser", "axId"]),
    (NS_C, "lineChart") => Set(["ser", "axId"]),
    (NS_C, "plotArea")  => Set(["areaChart","area3DChart","lineChart","line3DChart",
                                "stockChart","radarChart","scatterChart","pieChart",
                                "pie3DChart","doughnutChart","barChart","bar3DChart",
                                "ofPieChart","surfaceChart","surface3DChart","bubbleChart",
                                "valAx","catAx","dateAx","serAx"]),
    (NS_A, "p")          => Set(["r", "br", "fld"]),
    (NS_C, "legend")     => Set(["legendEntry"]),
)

# xsd:choice groups where every member excludes every other — at most one may
# appear. The common case.
const SCHEMA_ALTERNATIVES = Dict{Tuple{String,String},Vector{Vector{String}}}(
    (NS_A, "spPr") => [["custGeom","prstGeom"],
                       ["noFill","solidFill","gradFill","blipFill","pattFill","grpFill"],
                       ["effectLst","effectDag"]],
    (NS_C, "catAx")  => [["crosses","crossesAt"]],
    (NS_C, "valAx")  => [["crosses","crossesAt"]],
    (NS_C, "dateAx") => [["crosses","crossesAt"]],
    (NS_C, "serAx")  => [["crosses","crossesAt"]],
    (NS_C, "tx")     => [["strRef","rich"]],
    (NS_A, "ln")     => [["noFill","solidFill","gradFill","pattFill"],
                        ["prstDash","custDash"],
                        ["round","bevel","miter"]],
    (NS_A, "defRPr") => [["noFill","solidFill","gradFill","blipFill","pattFill","grpFill"],
                         ["effectLst","effectDag"],
                         ["uLnTx","uLn"],
                         ["uFillTx","uFill"]],
    (NS_A, "lnSpc")  => [["spcPct", "spcPts"]],
    (NS_A, "spcBef") => [["spcPct", "spcPts"]],
    (NS_A, "spcAft") => [["spcPct", "spcPts"]],
    (NS_A, "bodyPr") => [["noAutofit", "normAutofit", "spAutoFit"],
                         ["sp3d", "flatTx"]],
)
SCHEMA_ALTERNATIVES[(NS_A, "rPr")]        = SCHEMA_ALTERNATIVES[(NS_A, "defRPr")]
SCHEMA_ALTERNATIVES[(NS_A, "endParaRPr")] = SCHEMA_ALTERNATIVES[(NS_A, "defRPr")]

# Any member of a choice group identifies the group, since remove_choice removes
# whichever member is present rather than the one named. These constants say
# "the fill group" at a call site, where a bare "solidFill" would read as a target.
const FILL_GROUP   = "solidFill"
const JOIN_GROUP = "miter"
const EFFECT_GROUP = "effectLst"
const DASH_GROUP = "prstDash"
# c:tx is a choice of c:strRef or c:rich; either member names the group.
const TX_GROUP = "rich"

# xsd:choice where one member excludes a whole group whose own members coexist.
# In CT_DLbls, `delete` excludes every display property, but those properties
# appear together freely.
const SCHEMA_CHOICES = Dict{Tuple{String,String},Vector{Pair{String,Vector{String}}}}(
    (NS_C, "dLbls") => ["delete" => ["numFmt","spPr","txPr","dLblPos","showLegendKey",
                                     "showVal","showCatName","showSerName","showPercent",
                                     "showBubbleSize","separator","showLeaderLines",
                                     "leaderLines"]],
    (NS_C, "dLbl")  => ["delete" => ["layout","tx","numFmt","spPr","txPr","dLblPos",
                                     "showLegendKey","showVal","showCatName","showSerName",
                                     "showPercent","showBubbleSize","separator"]],
)

# Which series type governs a c:ser's child order, keyed by the group containing
# it. Not derivable from the group name: stockChart holds lineSer, doughnutChart
# and ofPieChart hold pieSer, and every 3D variant shares its 2D counterpart's
# series type.
const SER_TYPE = Dict{String,String}(
    "areaChart"      => "areaSer",
    "area3DChart"    => "areaSer",
    "lineChart"      => "lineSer",
    "line3DChart"    => "lineSer",
    "stockChart"     => "lineSer",
    "radarChart"     => "radarSer",
    "scatterChart"   => "scatterSer",
    "pieChart"       => "pieSer",
    "pie3DChart"     => "pieSer",
    "doughnutChart"  => "pieSer",
    "ofPieChart"     => "pieSer",
    "barChart"       => "barSer",
    "bar3DChart"     => "barSer",
    "surfaceChart"   => "surfaceSer",
    "surface3DChart" => "surfaceSer",
    "bubbleChart"    => "bubbleSer",
)

# The six c:dLbls display flags, in schema order. Excel writes all of them
# whenever it creates a c:dLbls.
const DLBLS_FLAGS = ("showLegendKey", "showVal", "showCatName", "showSerName",
                     "showPercent", "showBubbleSize")


# Excel UI names for a:prstDash values. Excel's dropdown offers eight of the
# eleven DrawingML presets; the other three (:dot, :sysDashDot, :sysDashDotDot)
# have no UI name and are reachable only by their DrawingML spelling.
#
# Note Round Dot is sysDot and Square Dot is sysDash — not :dot and :dash, which
# are different patterns Excel does not expose.
const DASH_ALIASES = Dict{Symbol,Symbol}(
    :roundDot        => :sysDot,
    :squareDot       => :sysDash,
    :longDash        => :lgDash,
    :longDashDot     => :lgDashDot,
    :longDashDotDot  => :lgDashDotDot,
    # :solid, :dash and :dashDot are spelled the same either way.
)

const CAP_ALIASES = Dict{Symbol,Symbol}(
    :square => :sq,
    :round  => :rnd,
    # :flat is the same either way.
)

const CMPD_ALIASES = Dict{Symbol,Symbol}(
    :simple    => :sng,
    :double    => :dbl,
    :triple    => :tri,
    # :thickThin and :thinThick are the same either way.
)

# ST_MarkerStyle (ECMA-376, dml-chart.xsd). Excel's UI uses the same words for
# the shapes, so there are no aliases; :auto and :picture have no UI entry.
const MARKER_SYMBOLS = (:circle, :dash, :diamond, :dot, :none, :picture,
                        :plus, :square, :star, :triangle, :x, :auto)


_with_children(n::XML.Node, kids::Vector) =
    typeof(n)(XML.nodetype(n), XML.tag(n), n.attributes, XML.value(n), kids)

# Slot index of `tag` in a child order, where a nested vector is one slot whose
# members may appear in any order (an unbounded xsd:choice). Nothing if unknown.
function _slot(order, tag::AbstractString)
    for (i, entry) in enumerate(order)
        entry isa AbstractString ? (entry == tag && return i) :
                                   (tag in entry && return i)
    end
    return nothing
end

function _schema_position(order, kids, pos::Integer)
    for (i, k) in enumerate(kids)
        p = _slot(order, localname(k))
        isnothing(p) && continue
        p > pos && return i
    end
    return length(kids) + 1
end

"""
    _check_choice(key, tag, kids)

Throw where inserting `tag` would violate an `xsd:choice` against a child already
present.
"""
function _check_choice(key::SchemaKey, tag::AbstractString, kids)
    for grp in get(SCHEMA_ALTERNATIVES, key, Vector{String}[])
        tag in grp || continue
        for k in kids
            kt = localname(k)
            if kt != tag && kt in grp
                throw(XLSXError(
                    "`$tag` and `$kt` are alternatives in `$(key[2])`; remove `$kt` first."))
            end
        end
    end
    for (owner, excluded) in get(SCHEMA_CHOICES, key, Pair{String,Vector{String}}[])
        conflicts = if tag == owner
            excluded
        elseif tag in excluded
            [owner]
        else
            continue
        end
        for k in kids
            kt = localname(k)
            if kt in conflicts
                throw(XLSXError(
                    "`$tag` and `$kt` are alternatives in `$(key[2])`; remove `$kt` first."))
            end
        end
    end
    return nothing
end

"""
    remove_choice(parent, key, member) -> XML.Node

Return a node equal to `parent` with whichever member of the `xsd:choice` group
containing `member` is present removed. Used to make a property inherit again:
removing the fill element from an `spPr` restores the cascade, which is not the
same as writing `<a:noFill/>`.
"""
function remove_choice(parent::XML.Node, key::SchemaKey, member::AbstractString)
    for grp in get(SCHEMA_ALTERNATIVES, key, Vector{String}[])
        member in grp || continue
        isnothing(parent.children) && return parent
        i = findfirst(k -> localname(k) in grp, parent.children)
        isnothing(i) && return parent
        kids = copy(parent.children)
        deleteat!(kids, i)
        return _with_children(parent, kids)
    end
    throw(XLSXError("`$member` is not part of a choice group in `$(key[2])`."))
end

"""
    insert_child(parent, key, child) -> XML.Node

Return a node equal to `parent` with `child` inserted at the position the schema
requires. Where the schema allows only one child of that tag and one is present,
it is replaced.

`key` is `parent`'s [`SchemaKey`](@ref) — `(NS_C, "barSer")` for a series in a
bar chart, `(NS_A, "spPr")` for shape properties anywhere.

`XML.Node` is immutable, so this returns the parent to use rather than mutating
it: `parent = insert_child(parent, key, child)`. Callers must splice the result
into its own parent — [`rebuild_path`](@ref) does that for a whole descent.

Throws where inserting `child` would violate an `xsd:choice` against a sibling
already present. Removing the loser is a decision for the setter, not a side
effect of insertion.
"""
function insert_child(parent::XML.Node, key::SchemaKey, child::XML.Node)
    order = get(CHILD_ORDER, key, nothing)
    isnothing(order) && throw(XLSXError("No child order known for `$key`."))

    tag = localname(child)
    pos = findfirst(==(tag), order)
    isnothing(pos) && throw(XLSXError("`$tag` is not a valid child of `$key`."))

    # `parent.children` is nothing for a childless element; XML.children flattens
    # that to an empty tuple, so only the field distinguishes the two cases.
    kids = isnothing(parent.children) ? XML.Node[] : copy(parent.children)

    _check_choice(key, tag, kids)

    if tag in get(REPEATABLE, key, Set{String}())
        last_same = findlast(k -> localname(k) == tag, kids)
        i = isnothing(last_same) ? _schema_position(order, kids, pos) : last_same + 1
        insert!(kids, i, child)
    else
        existing = findfirst(k -> localname(k) == tag, kids)
        if isnothing(existing)
            insert!(kids, _schema_position(order, kids, pos), child)
        else
            kids[existing] = child
        end
    end
    return _with_children(parent, kids)
end

"""
    replace_child(parent, old, new) -> XML.Node

Return a node equal to `parent` with the child identical (`===`) to `old`
replaced by `new`. Throws where `old` is not a child of `parent`.
"""
function replace_child(parent::XML.Node, old::XML.Node, new::XML.Node)
    isnothing(parent.children) &&
        throw(XLSXError("`$(localname(parent))` has no children."))
    i = findfirst(k -> k === old, parent.children)
    isnothing(i) && throw(XLSXError("Node is not a child of `$(localname(parent))`."))
    kids = copy(parent.children)
    kids[i] = new
    return _with_children(parent, kids)
end

"""
    remove_child(parent, tag) -> XML.Node

Return a node equal to `parent` with its child named `tag` removed, or `parent`
unchanged where it has none. Only the first match is removed.

`XML.Node` is immutable, so this returns the parent to use rather than mutating
it — see [`insert_child`](@ref).
"""
function remove_child(parent::XML.Node, tag::AbstractString)
    isnothing(parent.children) && return parent
    i = findfirst(k -> localname(k) == tag, parent.children)
    isnothing(i) && return parent
    kids = copy(parent.children)
    deleteat!(kids, i)
    return _with_children(parent, kids)
end

"""
    rebuild_path(node, steps, f; prefixes, parent_key) -> XML.Node

Descend `node` by `steps`, apply `f` to the element found there, and rebuild
every ancestor on the way back so the change appears in the returned root.

Each step is `key => tag` or `key => (tag, predicate)`, where `key` is the
[`SchemaKey`](@ref) of the element that step names. The key's namespace also
supplies the prefix for an element this creates.

`parent_key` is the [`SchemaKey`](@ref) of `node` itself, needed when a step has
to create an element and `insert_child` must know where it goes. It defaults to
the chart part's root; a caller that starts partway down must say so.

A step with no predicate that matches nothing is created empty and inserted in
schema order, so a descent always reaches a node. A step with a predicate is
never created — an arbitrary predicate does not say what would satisfy it — and
throws when nothing matches.

`f` takes the target element and returns its replacement.

Returns the new `node`. Every node from `node` to the target inclusive is a fresh
object; nodes off the path are shared. Any `raw` field held against a node on
that path is stale afterwards.

Never call this from a reader: it creates intermediates, so a getter built on it
would grow empty elements into the file just by looking at them.

# Example

    rebuild_path(chart_root(c),
                 [(NS_C, "chart")    => "chart",
                  (NS_C, "plotArea") => "plotArea",
                  (NS_C, "barChart") => "barChart",
                  (NS_C, "barSer")   => ("ser", n -> n === target),
                  (NS_A, "spPr")     => "spPr"],
                 sp -> insert_child(sp, (NS_A, "spPr"), fill_node);
                 prefixes = ns_prefixes(chart_root(c)))
"""
function rebuild_path(node::XML.Node, steps, f;
                      prefixes::Dict{String,String},
                      parent_key::SchemaKey = (NS_C, "chartSpace"))

    isempty(steps) && return f(node)

    step_key, spec = first(steps)
    tag, pred = spec isa Tuple ? spec : (spec, nothing)
    rest = steps[2:end]

    kids = isnothing(node.children) ? XML.Node[] : node.children
    i = findfirst(k -> localname(k) == tag && (isnothing(pred) || pred(k)), kids)

    if isnothing(i)
        isnothing(pred) ||
            throw(XLSXError("No `$tag` in `$(localname(node))` matches the predicate."))
        ns = step_key[1]
        haskey(prefixes, ns) ||
            throw(XLSXError("Namespace `$ns` is not declared on the chart part."))
        fresh = XML.Element(prefixed_tag(prefixes[ns], tag))
        node  = insert_child(node, parent_key, fresh)
        i     = findfirst(k -> k === fresh, node.children)
    end
    return replace_child(node, node.children[i],
                         rebuild_path(node.children[i], rest, f;
                                      prefixes, parent_key = step_key))
end

function _check(v::Symbol, allowed, what, aliases = Dict{Symbol,Symbol}())
    val = get(aliases, v, v)
    if val ∉ allowed     
        msg = "`$v` is not a valid $what. Valid values: " * join(allowed, ", ")
        isempty(aliases) || (msg *= "; or the Excel names " * join(keys(aliases), ", "))
        throw(XLSXError(msg * "."))
    end
    return val
end