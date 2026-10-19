"""
    getConditionalFormats(ws::Worksheet)

Get the conditional formats for a worksheet.

# Arguments
- `ws::Worksheet`: The worksheet for which to get the conditional formats.

Return a vector of pairs: CellRange => NamedTuple{type::String, priority::Int}}.


"""
getConditionalFormats(ws::Worksheet) = append!(getConditionalFormats(ws, allCfs(ws)), getConditionalExtFormats(ws, allExtCfs(ws)))
function getConditionalFormats(ws::Worksheet, allcfnodes::Vector{XML.Node})::Vector{Pair{CfKey,NamedTuple{(:type, :priority),Tuple{String,Int64}}}}
    allcfs = Vector{Pair{CfKey,NamedTuple{(:type, :priority),Tuple{String,Int64}}}}()
    wb = get_workbook(ws)
    for cf in allcfnodes
        for child in XML.children(cf)
            if localname(child) ==    "cfRule"
            push!(allcfs, _cf_range(ws, cf["sqref"]) => (type=child["type"], priority=parse(Int, child["priority"])))            end
        end
    end
    return allcfs
end
function getConditionalExtFormats(ws::Worksheet, allcfnodes::Vector{XML.Node})::Vector{Pair{CfKey,NamedTuple{(:type, :priority),Tuple{String,Int64}}}}
    wb = get_workbook(ws)
    allcfs = Vector{Pair{CfKey,NamedTuple{(:type, :priority),Tuple{String,Int64}}}}()
    for cf in allcfnodes
        let t, p, r, rule = false, ref = false
            localname(cf) == "conditionalFormatting" ||
                throw(XLSXError("Malformed worksheet `$(ws.name)`: its x14 conditional formats contain a `$(localname(cf))` element, not `conditionalFormatting`."))
            # Children may be interleaved with whitespace text nodes from the
            # writer's indentation, so locate `xm:sqref` by name rather than
            # assuming it is the last child.
            cf_elements = xml_elements(cf)
            sqref = findlast(c -> localname(c) == "sqref", cf_elements)
            if !isnothing(sqref)
                r = XML.simple_value(cf_elements[sqref])
                ref = true
            end
            for child in cf_elements
                if localname(child) == "cfRule"
                    t = child["type"]
                    if t != "dataBar" # This is the other half of a dataBar definition - don't count twice!
                        p = parse(Int, child["priority"])
                        rule = true
                    end
                end
                if rule && ref
                    push!(allcfs, _cf_range(ws, r) => (type=t, priority=p))
                    rule = false
                end
            end
        end
    end
    return allcfs
end
