define_cf_address_methods(:setCfColorScale)
function setCfColorScale(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int
    colorscale::Union{Nothing,String}=nothing
    min_type::Union{Nothing,String}="min"
    min_val::Union{Nothing,String}=nothing
    min_col::Union{Nothing,CellColor}="FFF8696B"
    mid_type::Union{Nothing,String}=nothing
    mid_val::Union{Nothing,String}=nothing
    mid_col::Union{Nothing,CellColor}="FFFCFCFF"
    max_type::Union{Nothing,String}="max"
    max_val::Union{Nothing,String}=nothing
    max_col::Union{Nothing,CellColor}="FFFFEB84"

    for (k, v) in allkws
        if k == :colorscale
            colorscale = v
        elseif k == :min_type
            min_type = v
        elseif k == :min_val
            min_val = v
        elseif k == :min_col
            min_col = v
        elseif k == :mid_type
            mid_type = v
        elseif k == :mid_val
            mid_val = v
        elseif k == :mid_col
            mid_col = v
        elseif k == :max_type
            max_type = v
        elseif k == :max_val
            max_val = v
        elseif k == :max_col
            max_col = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `colorscale`, `min_type`, `min_val`, `min_col`, `mid_type`, `mid_val`, `mid_col`, `max_type`, `max_val`, `max_col`."))
        end
    end

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"


    let new_pr, new_cf

        new_pr = string(next_cf_priority!(ws))

        if isnothing(colorscale)

            min_type == "min" && (min_val = nothing)
            max_type == "max" && (max_val = nothing)
            _check_cfvo("min", min_type, min_val, ("min", "percentile", "percent", "num", "formula"); optional=false)
            _check_cfvo("mid", mid_type, mid_val, ("percentile", "percent", "num", "formula"))
            _check_cfvo("max", max_type, max_val, ("max", "percentile", "percent", "num", "formula"); optional=false)

            min_val  = _normalise_cfvo_val(ws, min_val)
            mid_val  = _normalise_cfvo_val(ws, mid_val)
            max_val  = _normalise_cfvo_val(ws, max_val)

            cfx = XML.Element("$(pfx)cfRule"; type="colorScale", priority=new_pr)
            csc = XML.Element("$(pfx)colorScale")
            push!(csc,
                isnothing(min_val) ?
                    XML.Element("$(pfx)cfvo"; type=min_type) :
                    XML.Element("$(pfx)cfvo"; type=min_type, val=min_val)
            )

            !isnothing(mid_type) && push!(csc, XML.Element("$(pfx)cfvo"; type=mid_type, val=mid_val))

            push!(csc,
                isnothing(max_val) ?
                    XML.Element("$(pfx)cfvo"; type=max_type) :
                    XML.Element("$(pfx)cfvo"; type=max_type, val=max_val)
            )

            push!(csc, _color_element("$(pfx)color", min_col))
            !isnothing(mid_type) && push!(csc, _color_element("$(pfx)color", mid_col))
            push!(csc, _color_element("$(pfx)color", max_col))                
            push!(cfx, csc)
        else
            if !haskey(colorscales, colorscale)
                throw(XLSXError("Invalid colorscale option chosen: $colorscale. Valid options are: $(keys(colorscales))."))
            end
            cfx = copynode(colorscales[colorscale])
            cfx["priority"] = new_pr
        end

        update_worksheet_cfx!(cfx, ws, rng)

    end

    return 0
end
