"""
    clearConditionalFormats(ws::Worksheet, rng) -> Int

Remove every conditional format whose range falls entirely within `rng`, returning
the number of rules removed.

A rule is removed when all of the cells it applies to lie inside `rng`. A rule
spanning a wider range is left alone, even where it overlaps.

!!! note "Experimental"

    This function is experimental and is not exported. In particular the rule for which
    conditional formats are removed — currently, those lying entirely within rng — may
    change.

"""
function clearConditionalFormats end

clearConditionalFormats(ws::Worksheet, rng::AbstractString) =
    clearConditionalFormats(ws, CellRange(rng))

function clearConditionalFormats(ws::Worksheet, rng::CfRange)
    sheetdoc = xmlroot(get_workbook(ws), ws.relationship_id)
    n = 0

    # 2007 blocks: direct children of <worksheet>
    ws_el = sheetdoc[find_child_index(XML.children(sheetdoc), "worksheet")]
    for i in reverse(eachindex(XML.children(ws_el)))
        c = ws_el[i]
        localname(c) == "conditionalFormatting" || continue
        haskey(c, "sqref") || continue
        if _cf_within(ws, c["sqref"], rng)
            n += length(xml_elements(c))
            deleteat!(ws_el.children, i)
        end
    end

    # x14 counterparts under <extLst>
    blk = _x14_cf_block(sheetdoc)
    if !isnothing(blk)
        for q in reverse(eachindex(XML.children(blk)))
            b = blk[q]
            els = xml_elements(b)
            s = findlast(e -> localname(e) == "sqref", els)
            isnothing(s) && continue
            if _cf_within(ws, XML.simple_value(els[s]), rng)
                deleteat!(blk.children, q)
            end
        end
        # Excel refuses to open a file with an empty `<x14:conditionalFormattings/>`.
        if isempty(xml_elements(blk))
            i, j = get_idces(sheetdoc, "worksheet", "extLst")
            ext = _x14_cf_ext(sheetdoc[i][j])
            filter!(c -> c !== blk, ext.children)
        end
    end

    update_worksheets_xml!(get_xlsxfile(ws))
    return n
end

# Every area named by `sqref` lies inside `rng`.
function _cf_within(ws::Worksheet, sqref::AbstractString, rng::CfRange)
    inner = _cf_range(ws, sqref)
    outer = [a isa CellRef ? CellRange(a, a) : a for a in _cf_areas(rng)]
    for a in _cf_areas(inner)
        r = a isa CellRef ? CellRange(a, a) : a
        any(o -> issubset(r, o), outer) || return false
    end
    return true
end
