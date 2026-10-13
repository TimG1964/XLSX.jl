define_cf_address_methods(:setCfCellIs)
function setCfCellIs(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="greaterThan"
    value::Union{Nothing,String}=nothing
    value2::Union{Nothing,String}=nothing
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{<:Pair{String}}}=nothing
    font::Union{Nothing,Vector{<:Pair{String}}}=nothing
    border::Union{Nothing,Vector{<:Pair{String}}}=nothing
    fill::Union{Nothing,Vector{<:Pair{String}}}=nothing

    for (k, v) in allkws
        if k == :operator
            operator = v
        elseif k == :value
            value = v
        elseif k == :value2
            value2 = v
        elseif k == :stopIfTrue
            stopIfTrue = v
        elseif k == :dxStyle
            dxStyle = v
        elseif k == :format
            format = v
        elseif k == :font
            font = v
        elseif k == :border
            border = v
        elseif k == :fill
            fill = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid keywords are: `operator`, `value`, `value2`, `stopIfTrue`, `dxStyle`, `format`, `font`, `border` and `fill`."))
        end
    end

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"


    !isnothing(value) && !is_valid_cellname(value) && !is_valid_fixed_cellname(value) && isnothing(tryparse(Float64, value)) && throw(XLSXError("Invalid `value`: $value. Must be a number or a CellRef."))
    !isnothing(value2) && !is_valid_cellname(value2) && !is_valid_fixed_cellname(value2) && isnothing(tryparse(Float64, value2)) && throw(XLSXError("Invalid `value2`: $value2. Must be a number or a CellRef."))
    !isnothing(operator) && operator ∉ ["greaterThan", "greaterEqual", "lessThan", "lessEqual", "equal", "notEqual", "between", "notBetween"] && throw(XLSXError("Invalid `operator`: $operator. Valid options are: `greaterThan`, `greaterEqual`, `lessThan`, `lessEqual`, `equal`, `notEqual`, `between`, `notBetween`"))
    wb = get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    if isnothing(value)
        # The average of the range's numbers, as Excel's AVERAGE takes it: text,
        # booleans and empty cells are ignored.
        nums  = filter(v -> v isa Real && !(v isa Bool), _cf_values(ws, rng))
        value = isempty(nums) ? nothing : string(sum(nums) / length(nums))
    end

    cfx = XML.Element("$(pfx)cfRule"; type="cellIs", dxfId=string(dxid.id))
    cfx["priority"] = next_cf_priority!(ws)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["operator"] = operator
    push!(cfx, XML.Element("$(pfx)formula", XML.Text(value)))
    if !isnothing(value2) && operator ∈ ["between", "notBetween"]

        push!(cfx, XML.Element("$(pfx)formula", XML.Text(value2)))
    end

    update_worksheet_cfx!(cfx, ws, rng)

    return 0
end

define_cf_address_methods(:setCfContainsText)
function setCfContainsText(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="containsText"
    value::Union{Nothing,String}=nothing
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{<:Pair{String}}}=nothing
    font::Union{Nothing,Vector{<:Pair{String}}}=nothing
    border::Union{Nothing,Vector{<:Pair{String}}}=nothing
    fill::Union{Nothing,Vector{<:Pair{String}}}=nothing
    for (k, v) in allkws
        if k == :operator
            operator = String(v)
        elseif k == :value
            value = String(v)
        elseif k == :stopIfTrue
            stopIfTrue = String(v)
        elseif k == :dxStyle
            dxStyle = String(v)
        elseif k == :format
            format = v
        elseif k == :font
            font = v
        elseif k == :border
            border = v
        elseif k == :fill
            fill = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `operator`, `value`, `stopIfTrue`, `dxStyle`, `format`, `font`, `border`, `fill`."))
        end
    end

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"


    isnothing(value) && throw(XLSXError("Invalid `value`: $value. Must contain text or a CellRef."))

    wb = get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    type = operator
    if operator == "containsText"
        formula = "NOT(ISERROR(SEARCH(\"__txt__\",__CR__)))"
    elseif operator == "notContainsText"
        operator = "notContains"
        formula = "ISERROR(SEARCH(\"__txt__\",__CR__))"
    elseif operator == "beginsWith"
        #        operator = "beginsWith"
        formula = "LEFT(__CR__,LEN(\"__txt__\"))=\"__txt__\""
    elseif operator == "endsWith"
        #        operator = "endsWith"
        formula = "RIGHT(__CR__,LEN(\"__txt__\"))=\"__txt__\""
    else
        throw(XLSXError("Invalid operator: $type. Valid options are: `containsText`, `notContainsText`, `beginsWith`, `endsWith`."))
    end
    formula = replace(formula, "__txt__" => value, "__CR__" => string(_cf_anchor(rng)))

    cfx = XML.Element("$(pfx)cfRule"; type=type, dxfId=string(dxid.id))
    cfx["priority"] = next_cf_priority!(ws)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["operator"] = operator
    cfx["text"] = value
    push!(cfx, XML.Element("$(pfx)formula", XML.Text(formula)))

    update_worksheet_cfx!(cfx, ws, rng)

    return 0
end

define_cf_address_methods(:setCfTop10)
function setCfTop10(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="topN"
    value::Union{Nothing,String}="10"
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{<:Pair{String}}}=nothing
    font::Union{Nothing,Vector{<:Pair{String}}}=nothing
    border::Union{Nothing,Vector{<:Pair{String}}}=nothing
    fill::Union{Nothing,Vector{<:Pair{String}}}=nothing
    for (k, v) in allkws
        if k == :operator
            operator = v
        elseif k == :value
            value = v
        elseif k == :stopIfTrue
            stopIfTrue = v
        elseif k == :dxStyle
            dxStyle = v
        elseif k == :format
            format = v
        elseif k == :font
            font = v
        elseif k == :border
            border = v
        elseif k == :fill
            fill = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `operator`, `value`, `stopIfTrue`, `dxStyle`, `format`, `font`, `border`, `fill`."))
        end
    end

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"


    !isnothing(value) && !is_valid_cellname(value) && !is_valid_fixed_cellname(value) && isnothing(tryparse(Float64, value)) && throw(XLSXError("Invalid `value`: $value. Must be a number or a CellRef."))

    wb = get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    percent = ""
    bottom = ""
    cfx = XML.Element("$(pfx)cfRule"; type="top10", dxfId=string(dxid.id))
    if operator == "topN"
    elseif operator == "topN%"
        percent = "1"
    elseif operator == "bottomN"
        bottom = "1"
    elseif operator == "bottomN%"
        percent = "1"
        bottom = "1"
    else
        throw(XLSXError("Invalid operator: $operator. Valid options are: `topN`, `topN%`, `bottomN`, `bottomN%`."))
    end

    cfx["priority"] = next_cf_priority!(ws)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    if percent != ""
        cfx["percent"] = percent
    end
    if bottom != ""
        cfx["bottom"] = bottom
    end
    cfx["rank"] = value

    update_worksheet_cfx!(cfx, ws, rng)

    return 0
end

define_cf_address_methods(:setCfAboveAverage)
function setCfAboveAverage(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="aboveAverage"
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{<:Pair{String}}}=nothing
    font::Union{Nothing,Vector{<:Pair{String}}}=nothing
    border::Union{Nothing,Vector{<:Pair{String}}}=nothing
    fill::Union{Nothing,Vector{<:Pair{String}}}=nothing
    for (k, v) in allkws
        if k == :operator
            operator = v
        elseif k == :stopIfTrue
            stopIfTrue = v
        elseif k == :dxStyle
            dxStyle = v
        elseif k == :format
            format = v
        elseif k == :font
            font = v
        elseif k == :border
            border = v
        elseif k == :fill
            fill = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `operator`, `stopIfTrue`, `dxStyle`, `format`, `font`, `border`, `fill`."))
        end
    end

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"


    wb = get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    if operator == "aboveAverage"
        cfx = XML.Element("$(pfx)cfRule"; type=operator, dxfId=string(dxid.id), priority="1")
    elseif operator == "aboveEqAverage"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", equalAverage="1")
    elseif operator == "plus1StdDev"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", stdDev="1")
    elseif operator == "plus2StdDev"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", stdDev="2")
    elseif operator == "plus3StdDev"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", stdDev="3")
    elseif operator == "belowAverage"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", aboveAverage="0")
    elseif operator == "belowEqAverage"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", aboveAverage="0", equalAverage="1")
    elseif operator == "minus1StdDev"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", aboveAverage="0", stdDev="1")
    elseif operator == "minus2StdDev"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", aboveAverage="0", stdDev="2")
    elseif operator == "minus3StdDev"
        cfx = XML.Element("$(pfx)cfRule"; type="aboveAverage", dxfId=string(dxid.id), priority="1", aboveAverage="0", stdDev="3")
    else
        throw(XLSXError("Invalid operator: $operator. Valid options are: `aboveAverage`, `aboveEqAverage`, `plus1sStdDev`, `plus2StdDev`, `plus3StdDev`, `belowAverage`, `belowEqAverage`, `minus1StdDev`, `minus2StdDev`, `minus3StdDev`."))
    end

    cfx["priority"] = next_cf_priority!(ws)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end

    update_worksheet_cfx!(cfx, ws, rng)

    return 0
end

define_cf_address_methods(:setCfTimePeriod)
function setCfTimePeriod(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="last7Days"
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{<:Pair{String}}}=nothing
    font::Union{Nothing,Vector{<:Pair{String}}}=nothing
    border::Union{Nothing,Vector{<:Pair{String}}}=nothing
    fill::Union{Nothing,Vector{<:Pair{String}}}=nothing
    for (k, v) in allkws
        if k == :operator
            operator = v
        elseif k == :stopIfTrue
            stopIfTrue = v
        elseif k == :dxStyle
            dxStyle = v
        elseif k == :format
            format = v
        elseif k == :font
            font = v
        elseif k == :border
            border = v
        elseif k == :fill
            fill = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `operator`, `stopIfTrue`, `dxStyle`, `format`, `font`, `border`, `fill`."))
        end
    end

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"


    if haskey(timeperiods, operator)
        formula = timeperiods[operator]
    else
        throw(XLSXError("Invalid operator: $operator. Valid options are: `yesterday`, `today`, `tomorrow`, `last7Days`, `lastWeek`, `thisWeek`, `nextWeek`, `lastMonth`, `thisMonth`, `nextMonth`."))
    end

    formula = replace(formula, "__CR__" => string(_cf_anchor(rng)))

    wb = get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    cfx = XML.Element("$(pfx)cfRule"; type="timePeriod", dxfId=string(dxid.id))
    cfx["priority"] = next_cf_priority!(ws)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["timePeriod"] = operator

    push!(cfx, XML.Element("$(pfx)formula", XML.Text(formula)))

    update_worksheet_cfx!(cfx, ws, rng)

    return 0
end

define_cf_address_methods(:setCfContainsBlankErrorUniqDup)
function setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="containsBlanks"
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{<:Pair{String}}}=nothing
    font::Union{Nothing,Vector{<:Pair{String}}}=nothing
    border::Union{Nothing,Vector{<:Pair{String}}}=nothing
    fill::Union{Nothing,Vector{<:Pair{String}}}=nothing
    for (k, v) in allkws
        if k == :operator
            operator = v
        elseif k == :stopIfTrue
            stopIfTrue = v
        elseif k == :dxStyle
            dxStyle = v
        elseif k == :format
            format = v
        elseif k == :font
            font = v
        elseif k == :border
            border = v
        elseif k == :fill
            fill = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `operator`, `stopIfTrue`, `dxStyle`, `format`, `font`, `border`, `fill`."))
        end
    end

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"


    if operator == "containsBlanks"
        formula = "LEN(TRIM(__CR__))=0"
    elseif operator == "notContainsBlanks"
        formula = "LEN(TRIM(__CR__))>0"
    elseif operator == "containsErrors"
        formula = "ISERROR(__CR__)"
    elseif operator == "notContainsErrors"
        formula = "NOT(ISERROR(__CR__))"
    elseif operator == "uniqueValues"
        formula = ""
    elseif operator == "duplicateValues"
        formula = ""
    end
    formula = replace(formula, "__CR__" => string(_cf_anchor(rng)))

    wb = get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    cfx = XML.Element("$(pfx)cfRule"; type=operator, dxfId=string(dxid.id))
    cfx["priority"] = next_cf_priority!(ws)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    formula != "" && push!(cfx, XML.Element("$(pfx)formula", XML.Text(formula)))

    update_worksheet_cfx!(cfx, ws, rng)

    return 0
end

define_cf_address_methods(:setCfFormula)
function setCfFormula(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    formula::Union{Nothing,String}=nothing
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{<:Pair{String}}}=nothing
    font::Union{Nothing,Vector{<:Pair{String}}}=nothing
    border::Union{Nothing,Vector{<:Pair{String}}}=nothing
    fill::Union{Nothing,Vector{<:Pair{String}}}=nothing
    for (k, v) in allkws
        if k == :formula
            formula = v
        elseif k == :stopIfTrue
            stopIfTrue = v
        elseif k == :dxStyle
            dxStyle = v
        elseif k == :format
            format = v
        elseif k == :font
            font = v
        elseif k == :border
            border = v
        elseif k == :fill
            fill = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `formula`, `stopIfTrue`, `dxStyle`, `format`, `font`, `border`, `fill`."))
        end
    end
    isnothing(formula) && throw(XLSXError("A `formula` must be provided as a keyword argument."))

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"


    wb = get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    cfx = XML.Element("$(pfx)cfRule"; type="expression", dxfId=string(dxid.id))
    cfx["priority"] = next_cf_priority!(ws)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end

    push!(cfx, XML.Element("$(pfx)formula", XML.Text("(" * uppercase_unquoted(formula) * ")")))

    update_worksheet_cfx!(cfx, ws, rng)

    return 0
end
