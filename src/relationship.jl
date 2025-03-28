
function Relationship(e::XML.Node) :: Relationship
    XML.tag(e) != "Relationship" && throw(XLSXError("Unexpected XMLElement: $(XML.tag(e)). Expected: \"Relationship\"."))
    a = XML.attributes(e)
    return Relationship(
        a["Id"],
        a["Type"],
        a["Target"]
    )
end

function parse_relationship_target(prefix::String, target::String) :: String
    isempty(prefix) || isempty(target) && throw(XLSXError("Something wrong here!"))

    if target[1] == '/'
        sizeof(target) <= 1 && throw(XLSXError("Incomplete target path $target."))
        return target[2:end]
    else
        return prefix * '/' * target
    end
end

function get_relationship_target_by_id(prefix::String, wb::Workbook, Id::String) :: String
    for r in wb.relationships
        if Id == r.Id
            return parse_relationship_target(prefix, r.Target)
        end
    end
    error("Relationship Id=$(Id) not found")
end

function get_relationship_target_by_type(prefix::String, wb::Workbook, _type_::String) :: String
    for r in wb.relationships
        if _type_ == r.Type
            return parse_relationship_target(prefix, r.Target)
        end
    end
    error("Relationship Type=$(_type_) not found")
end

function has_relationship_by_type(wb::Workbook, _type_::String) :: Bool
    for r in wb.relationships
        if _type_ == r.Type
            return true
        end
    end
    false
end

function get_package_relationship_root(xf::XLSXFile) :: XML.Node
    xroot = xmlroot(xf, "_rels/.rels")[end]
    XML.tag(xroot) != "Relationships" && throw(XLSXError("Malformed XLSX file $(xf.source). _rels/.rels root node name should be `Relationships`. Found $(XML.tag(xroot))."))
    if (""=>"http://schemas.openxmlformats.org/package/2006/relationships") ∉ get_namespaces(xroot) 
        throw(XLSXError("Unexpected namespace at workbook relationship file: `$(get_namespaces(xroot))`."))
    end
    return xroot
end

function get_workbook_relationship_root(xf::XLSXFile) :: XML.Node
    xroot = xmlroot(xf, "xl/_rels/workbook.xml.rels")[end]
    XML.tag(xroot) != "Relationships" && throw(XLSXError("Malformed XLSX file $(xf.source). xl/_rels/workbook.xml.rels root node name should be `Relationships`. Found $(XML.tag(xroot))."))
    if (""=>"http://schemas.openxmlformats.org/package/2006/relationships") ∉ get_namespaces(xroot)
        throw(XLSXError("Unexpected namespace at workbook relationship file: `$(get_namespaces(xroot))`."))
    end
    return xroot
end

# Adds new relationship. Returns new generated rId.
function add_relationship!(wb::Workbook, target::String, _type::String) :: String
    xf = get_xlsxfile(wb)
    !is_writable(xf) && throws(XLSXError("XLSXFile instance is not writable."))
    local rId :: String

    let
        got_unique_id = false
        id = 1

        while !got_unique_id
            got_unique_id = true
            rId = string("rId", id)
            for r in wb.relationships
                if r.Id == rId
                    got_unique_id = false
                    id += 1
                    break
                end
            end
        end
    end

    # adds to relationship vector
    new_relationship = Relationship(rId, _type, target)
    push!(wb.relationships, new_relationship)

    # adds to XML tree
    xroot = get_workbook_relationship_root(xf)
    el = XML.Element("Relationship"; Id = rId, Target = target, Type = _type)
    push!(xroot, el)

    return rId
end
