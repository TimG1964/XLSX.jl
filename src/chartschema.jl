# Child order for xsd:sequence elements in the c: chart schema (ECMA-376,
# dml-chart.xsd), denormalized — EG_SerShared and EG_AxShared expanded inline.
# An element inserted out of this order makes the part invalid and Excel will
# refuse or repair the file.
const CHILD_ORDER = Dict{Tuple{String,String},Vector{String}}(

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
    (NS_A, "spPr") => ["xfrm","noFill","solidFill","gradFill","blipFill","pattFill",
                       "grpFill","ln","effectLst","effectDag","scene3d","sp3d","extLst"],
    (NS_A, "txPr") => ["bodyPr","lstStyle","p"],
    (NS_A, "rich") => ["bodyPr","lstStyle","p"],

    # CT_Tx is a chart type, unlike the two above.
    (NS_C, "tx")   => ["strRef","rich"],
)

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
)

# xsd:choice groups where every member excludes every other — at most one may
# appear. The common case.
const SCHEMA_ALTERNATIVES = Dict{Tuple{String,String},Vector{Vector{String}}}(
    (NS_A, "spPr") => [["noFill","solidFill","gradFill","blipFill","pattFill","grpFill"],
                       ["effectLst","effectDag"]],
    (NS_C, "catAx")  => [["crosses","crossesAt"]],
    (NS_C, "valAx")  => [["crosses","crossesAt"]],
    (NS_C, "dateAx") => [["crosses","crossesAt"]],
    (NS_C, "serAx")  => [["crosses","crossesAt"]],
    (NS_C, "tx")     => [["strRef","rich"]],
)

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


"""
    SchemaKey

The `(namespace, complex-type)` pair identifying which `xsd:sequence` governs an
element's children. Not derivable from the element's tag: every series is `c:ser`
but its child order depends on the group containing it (`barSer`, `lineSer`, …),
and `c:spPr` is `a:CT_ShapeProperties` despite its chart prefix. Callers state it.
"""
const SchemaKey = Tuple{String,String}

_with_children(n::XML.Node, kids::Vector) =
    typeof(n)(XML.nodetype(n), XML.tag(n), n.attributes, XML.value(n), kids)

# First index whose tag sorts after `pos` in the schema order; end+1 if none.
function _schema_position(order, kids, pos::Integer)
    for (i, k) in enumerate(kids)
        p = findfirst(==(localname(k)), order)
        isnothing(p) && continue           # unknown child (extension) — leave it alone
        p > pos && return i
    end
    return length(kids) + 1
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
    rebuild_path(node, steps, f; prefixes, key) -> XML.Node

Descend `node` by `steps`, apply `f` to the element found there, and rebuild
every ancestor on the way back so the change appears in the returned root.

Each step is `key => tag` or `key => (tag, predicate)`, where `key` is the
[`SchemaKey`](@ref) of the element that step names. The key's namespace also
supplies the prefix for an element this creates. `key` (the keyword) is the
schema key of `node` itself, defaulting to the chart part's root.

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
                      key::SchemaKey = (NS_C, "chartSpace"))
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
        node  = insert_child(node, key, fresh)
        i     = findfirst(k -> k === fresh, node.children)
    end
    return replace_child(node, node.children[i],
                         rebuild_path(node.children[i], rest, f;
                                      prefixes, key = step_key))
end
