
@inline Base.isempty(::EmptyCell) = true
@inline Base.isempty(::AbstractCell) = false
@inline iserror(c::Cell) = c.datatype == CT_ERROR
@inline iserror(::AbstractCell) = false
@inline row_number(c::EmptyCell) = row_number(c.ref)
@inline column_number(c::EmptyCell) = column_number(c.ref)
@inline row_number(c::Cell) = row_number(c.ref)
@inline column_number(c::Cell) = column_number(c.ref)
@inline relative_cell_position(c::Cell, rng::CellRange) = relative_cell_position(c.ref, rng)
@inline relative_cell_position(c::EmptyCell, rng::CellRange) = relative_cell_position(c.ref, rng)
@inline relative_column_position(c::Cell, rng::ColumnRange) = relative_column_position(c.ref, rng)
@inline relative_column_position(c::EmptyCell, rng::ColumnRange) = relative_column_position(c.ref, rng)

Base.:(==)(c1::Cell, c2::Cell) = c1.ref == c2.ref && c1.datatype == c2.datatype && c1.style == c2.style && c1.value == c2.value && c1.meta == c2.meta && c1.formula == c2.formula
Base.hash(c::Cell) = hash(c.ref) + hash(c.datatype) + hash(c.style) + hash(c.value) + hash(c.meta) + hash(c.formula)

Base.:(==)(c1::EmptyCell, c2::EmptyCell) = c1.ref == c2.ref
Base.hash(c::EmptyCell) = hash(c.ref) + 10

const RGX_INTEGER = r"^\-?[0-9]+$"
function get_error_type(v::AbstractString)::UInt64
    if v == "#NULL!"
        return UInt64(XL_NULL)
    elseif v == "#DIV/0!"
        return UInt64(XL_DIV0)
    elseif v == "#VALUE!"
        return UInt64(XL_VALUE)
    elseif v == "#REF!"
        return UInt64(XL_REF)
    elseif v == "#NAME?"
        return UInt64(XL_NAME)
    elseif v == "#NUM!"
        return UInt64(XL_NUM)
    elseif v == "#N/A"
        return UInt64(XL_NA)
    elseif v == "#SPILL!"
        return UInt64(XL_SPILL)
    else
        throw(XLSXError("Unknown error value: $v"))
    end
end
function get_error_string(e::UInt64)::String
    if e == UInt64(CellErrorType.XL_NULL)
        return "#NULL!"
    elseif e == UInt64(CellErrorType.XL_DIV0)
        return "#DIV/0!"
    elseif e == UInt64(CellErrorType.XL_VALUE)
        return "#VALUE!"
    elseif e == UInt64(CellErrorType.XL_REF)
        return "#REF!"
    elseif e == UInt64(CellErrorType.XL_NAME)
        return "#NAME?"
    elseif e == UInt64(CellErrorType.XL_NUM)
        return "#NUM!"
    elseif e == UInt64(CellErrorType.XL_NA)
        return "#N/A"
    elseif e == UInt64(CellErrorType.XL_SPILL)
        return "#SPILL!"
     else
        throw(XLSXError("Unknown error code: $e"))
    end
end

function Cell(c::XML.LazyNode, ws::Worksheet; mylock::Union{ReentrantLock,Nothing}=nothing)::Union{Cell,EmptyCell}
    wb = get_workbook(ws)

    # Validate tag first (fail fast)
    XML.tag(c) == "c" || throw(XLSXError("`Cell` Expects a `c` (cell) XML node."))

    a = XML.attributes(c)
    chn = XML.children(c)
    ref = CellRef(a["r"])

    # Get attributes once
    t = get(a, "t", "")
    s_str = get(a, "s", "")
    m_str = get(a, "cm", "")

    # Pre-allocate with concrete types
    datatype::CellValueType = CT_EMPTY
    style::UInt32 = isempty(s_str) ? UInt32(0) : parse(UInt32, s_str)
    value::UInt64 = UInt64(0)
    meta::UInt32 = isempty(m_str) ? UInt32(0) : parse(UInt32, m_str) + UInt32(1)
    formula::Bool = false

    if t == "inlineStr"
        # Handle inlineStr case - find "is" element
        for child in chn
            XML.tag(child) == "is" || continue

            uft = unformatted_text(child)
            if !isempty(uft)
                # Build formatted text - use smaller initial buffer
                io = IOBuffer()
                write(io, "<si>\n  ")

                # Write children more efficiently
                children_list = XML.children(child)
                n = length(children_list)
                for i in 1:n
                    i > 1 && write(io, "\n")
                    write(io, XML.write(children_list[i]))
                end

                write(io, "\n</si>")
                ft = String(take!(io))

                datatype = CT_STRING
                value = reinterpret(UInt64, Int64(add_formatted_string!(wb, ft; mylock)))
            end
            break
        end
    else
        # Parse style number once if needed
        num_style = isempty(s_str) ? 0 : parse(Int, s_str)
        
        for child in chn
            tag = XML.tag(child)
            
            if tag == "v"
                ch = XML.children(child)
                isempty(ch) && continue
                
                v = XML.unescape(XML.value(ch[1]))
                datatype, value = process_tv(wb, t, v, num_style; mylock)
            elseif tag == "f"
                if get_xlsxfile(wb).is_writable # only store formulas when XLSXFile is writable
                    f = parse_formula_from_element(child)
                    wb.formulas[SheetCellRef(combine_sheet_ref(ws, ref))] = f
                end
                formula = true
            end
        end
    end
    return Cell(ref, value, style, meta, datatype, formula)
end

function parse_formula_from_element(c_child_element)::AbstractFormula

    if XML.tag(c_child_element) != "f"
        throw(XLSXError("Expected nodename `f`. Found: `$(XML.tag(c_child_element))`"))
    end

    if XML.is_simple(c_child_element)
        formula_string = XML.unescape(XML.simple_value(c_child_element))
    else
        fs = [x for x in XML.eachchild(c_child_element) if XML.nodetype(x) == XML.Text]
        if length(fs) == 0
            formula_string = ""
        else
            formula_string = XML.unescape(XML.value(fs[1]))
        end
    end

    a = XML.attributes(c_child_element)
    unhandled_attributes = Dict{String,String}()
    if !isnothing(a)
        for (k, v) in a
            if k ∉ ["t", "si", "ref"]
                push!(unhandled_attributes, k => v)
            end
        end
    end
    is_array = false
    let ref = nothing
        if !isnothing(a)
            if haskey(a, "t")
                if a["t"] == "shared"
                    haskey(a, "si") || throw(XLSXError("Expected shared formula to have an index. `si` attribute is missing: $c_child_element"))
                    if haskey(a, "ref")
                        return ReferencedFormula(
                            formula_string,
                            parse(Int, a["si"]),
                            a["ref"],
                            length(unhandled_attributes) > 0 ? unhandled_attributes : nothing,
                        )
                    else
                        return FormulaReference(
                            parse(Int, a["si"]),
                            length(unhandled_attributes) > 0 ? unhandled_attributes : nothing,
                        )
                    end
                elseif a["t"] == "array"
                    is_array = true
                    ref = haskey(a, "ref") ? a["ref"] : nothing
                end
            end
        end
        return Formula(
            formula_string,
            is_array ? "array" : nothing,
            ref,
            length(unhandled_attributes) > 0 ? unhandled_attributes : nothing)
    end
end

function process_tv(wb::Workbook, t::String, v::String, num_style::Int; mylock::Union{ReentrantLock,Nothing}=nothing)
    datatype::CellValueType = CT_EMPTY
    value::UInt64 = UInt64(0)
    v == "" && (return datatype, value)

    if t == "b"
        # Boolean - avoid branches
        datatype = CT_BOOL
        value = v == "1" ? UInt64(1) : (v == "0" ? UInt64(0) : throw(XLSXError("Unknown boolean value: $v")))
        
    elseif t == "s"
        # Shared String
        datatype = CT_STRING
        value = reinterpret(UInt64, parse(Int64, v))
        
    elseif t == "str"
        # Plain String
        datatype = CT_STRING
        value = reinterpret(UInt64, Int64(add_shared_string!(wb, v; mylock)))

    elseif t == "e"
        # Error
        datatype = CT_ERROR
        value = get_error_type(v)
        
    elseif t == "n" || t == ""
        # Number - check datetime/float style once
        if styles_is_datetime(wb, num_style)
            # dates & times
            has_decimal = occursin('.', v)
            if has_decimal || v == "0"
                time_value = parse(Float64, v)
                time_value < 0 && throw(XLSXError("Cannot have a datetime value < 0. Got $time_value"))
                value = reinterpret(UInt64, time_value)
                if time_value <= 1.0
                    datatype = CellValueType(6)
                else
                    datatype = CellValueType(7)
                end
            else
                # Date
                time_value = parse(Int64, v)
                time_value < 0 && throw(XLSXError("Cannot have a datetime value < 0. Got $time_value"))
                value = reinterpret(UInt64, time_value)
                datatype = CT_DATE
            end
        elseif styles_is_float(wb, num_style)
            # float
            datatype = CT_FLOAT
            value = reinterpret(UInt64, parse(Float64, v))
        else
            # Check if integer using character-by-character scan
            is_int = true
            for i in 1:ncodeunits(v)
                c = codeunit(v, i)
                if !((c >= 0x30 && c <= 0x39) || (i == 1 && (c == 0x2d || c == 0x2b)))
                    # Not 0-9, or not leading +/-
                    is_int = false
                    break
                end
            end
            
            if is_int && !isempty(v)
                datatype = CT_INT
                value = reinterpret(UInt64, parse(Int64, v))
            else
                if ismissing(v) || isempty(v)
                    datatype=CT_EMPTY
                    value=UInt64(0)
                else
                    datatype = CT_FLOAT
                    value = reinterpret(UInt64, parse(Float64, v))
                end
            end
        end
    else
        throw(XLSXError("Cannot parse cell value: $v"))
    end

    return datatype, value
end

# Constructor with simple formula string for backward compatibility & tests
function Cell(wb::Workbook, ref::CellRef, t::String, s::String, v::String, m::String, f::Bool)
    style::UInt32 = isempty(s) ? UInt32(0) : parse(UInt32, s)
    meta::UInt32 = isempty(m) ? UInt32(0) : parse(UInt32, m) + UInt32(1)

    num_style = isempty(s) ? 0 : parse(Int, s)
    datatype, value = process_tv(wb, t, v, num_style)

    return Cell(ref, value, style, meta, datatype, f)
end
#function Cell(ref::CellRef, datatype::String, style::String, value::String, meta::String, formula::String)
#    println("What am I doing here?")
#    if formula == ""
#        return Cell(ref, datatype, style, value, meta, Formula())
#    else
#        return Cell(ref, datatype, style, value, meta, Formula(formula))
#    end
#end

@inline getdata(ws::Worksheet, empty::EmptyCell) = missing

"""
    getdata(ws::Worksheet, cell::Cell) :: CellValue

Returns a Julia representation of a given cell value.
The result data type is chosen based on the value of the cell as well as its style.

For example, date is stored as integers inside the spreadsheet, and the style is the
information that is taken into account to chose `Date` as the result type.

For numbers, if the style implies that the number is visualized with decimals,
the method will return a float, even if the underlying number is stored
as an integer inside the spreadsheet XML.

If `cell` has empty value or empty `String`, this function will return `missing`.
"""
function getdata(ws::Worksheet, cell::Cell)
    is1904 = isdate1904(get_workbook(ws))
    
    if iserror(cell)
        return missing 
    end

    if cell.datatype == CT_EMPTY
        return missing
    end

    if cell.datatype == CT_STRING
        # use sst
        str = sst_unformatted_string(ws, reinterpret(Int64, cell.value))
        return str

    elseif cell.datatype == CT_DATETIME
        # datetime
        return excel_value_to_datetime(reinterpret(Float64, cell.value), is1904)

    elseif cell.datatype == CT_DATE
        # datetime
        return excel_value_to_date(reinterpret(Int64, cell.value), is1904)

    elseif cell.datatype == CT_TIME
        # datetime
        return excel_value_to_time(reinterpret(Float64, cell.value))

    elseif cell.datatype == CT_BOOL
        # boolean
        return cell.value != 0

    elseif cell.datatype == CT_FLOAT
        # float
        return reinterpret(Float64, cell.value)

    elseif cell.datatype == CT_INT
        # int
        return reinterpret(Int64, cell.value)

    elseif cell.datatype == CT_ERROR
        # Error
        return missing
    end

    throw(XLSXError("Couldn't parse data for $cell."))
end

function _celldata_datetime(v::AbstractString, _is_date_1904::Bool)# :: Union{Dates.DateTime, Dates.Date, Dates.Time}

    # does not allow empty string
    if isempty(v)
        throw(XLSXError("Cannot convert an empty string into a datetime value."))
    end

    if occursin(".", v) || v == "0"
        time_value = parse(Float64, v)
        if time_value < 0
            throw(XLSXError("Cannot have a datetime value < 0. Got $time_value"))
        end

        if time_value <= 1
            # Time
            return excel_value_to_time(time_value), CellValueType.CT_TIME
        else
            # DateTime
            return excel_value_to_datetime(time_value, _is_date_1904), CellValueType.CT_DATETIME
        end
    else
        # Date
        return excel_value_to_date(parse(Int, v), _is_date_1904), CellValueType.CT_CT_DATE
    end
end

# Converts Excel number to Time.
# `x` must be between 0 and 1.
# To represent Time, Excel uses the decimal part
# of a floating point number. `1` equals one day.
function excel_value_to_time(x::Float64)::Dates.Time
    if x >= 0 && x <= 1
        return Dates.Time(Dates.Nanosecond(round(Int, x * 86400) * 1E9))
    else
        throw(XLSXError("A value must be between 0 and 1 to be converted to time. Got $x"))
    end
end

time_to_excel_value(x::Dates.Time)::Float64 = Dates.value(x) / (86400 * 1E9)

# Converts Excel number to Date. See also XLSX.isdate1904.
function excel_value_to_date(x::Int, _is_date_1904::Bool)::Dates.Date
    if _is_date_1904
        return Dates.Date(Dates.rata2datetime(x + 695056))
    else
        return Dates.Date(Dates.rata2datetime(x + 693594))
    end
end

function date_to_excel_value(date::Dates.Date, _is_date_1904::Bool)::Int
    if _is_date_1904
        return Dates.datetime2rata(date) - 695056
    else
        return Dates.datetime2rata(date) - 693594
    end
end

# Converts Excel number to DateTime.
# The decimal part represents the Time (see `_time` function).
# The integer part represents the Date.
# See also XLSX.isdate1904.
function excel_value_to_datetime(x::Float64, _is_date_1904::Bool)::Dates.DateTime
    if x < 0
        throw(XLSXError("Cannot have a datetime value < 0. Got $x"))
    end

    local dt::Dates.Date
    local hr::Dates.Time

    dt_part = trunc(Int, x)
    hr_part = x - dt_part

    dt = excel_value_to_date(dt_part, _is_date_1904)
    hr = excel_value_to_time(hr_part)

    return dt + hr
end

function datetime_to_excel_value(dt::Dates.DateTime, _is_date_1904::Bool)::Float64
    return date_to_excel_value(Dates.Date(dt), _is_date_1904) + time_to_excel_value(Dates.Time(dt))
end

# Extract cells from a <row> LazyNode and push them (in place) into a Dict(column -> Cell)
function get_rowcells!(rowcells::Dict{Int,Cell}, row::XML.LazyNode, ws::Worksheet; mylock::Union{ReentrantLock,Nothing}=nothing)

    #=
        # threaded cell extraction causes hugely more lock conflicts for low cellchunk size.
        # may be worthwhile if many columns (hundreds+), with a cellchunk size > ~10 or ~20, but this is unverified.

        # debug
        # @assert row.tag == "row" "Not a row node"
        cellchunk=8 # bigger chunks, fewer lock conflicts but columns are generally relatively few.
        sst_count=0
        d=row.depth

        row_cellnodes = Channel{Vector{XML.LazyNode}}(1 << 10)
        row_cells = Channel{Vector{XLSX.Cell}}(1 << 10)

        # consumer task
        consumer = @async begin
            for cells in row_cells  
                for cell in cells      
                    sst_count += cell.datatype == "s" ? 1 : 0
                    rowcells[column_number(cell)] = cell
                end
            end
        end

        # Feed row_cellnodes
        cellnodes = Vector{XML.LazyNode}(undef, cellchunk)
        pos=0
        cellnode=XML.next(row)
        while !isnothing(cellnode) && cellnode.depth > d
            if cellnode.tag == "c" # This is a cell
                pos += 1
                cellnodes[pos] = cellnode
            end
            if pos >= cellchunk
                put!(row_cellnodes, copy(cellnodes))
                pos=0
            end
            cellnode = XML.next(cellnode)
        end
        if pos>0 # handle last incomplete chunk
            put!(row_cellnodes, cellnodes[1:pos])
        end
        close(row_cellnodes)

        # Producer tasks
        mylock = ReentrantLock() # lock for thread-safe access to shared string table in case of inlineStrings
        @sync for _ in 1:Threads.nthreads()
            Threads.@spawn begin
                chunk = Vector{XLSX.Cell}(undef, cellchunk)
                for cns in row_cellnodes
                    cell_count=0
                    for cn in cns
                        cell_count += 1
                        chunk[cell_count] = Cell(cn, ws; mylock)
                        if cell_count >= cellchunk
                            put!(row_cells, copy(chunk))
                            cell_count=0
                        end
                    end
                    if cell_count > 0 # handle last incomplete chunk
                        put!(row_cells, chunk[1:cell_count])
                    end
                end
            end
        end
        close(row_cells)

        wait(consumer)  # ensure consumer is done

        if !isnothing(cellnode) && cellnode.tag == "row" # have reached the end of last row, beginning of next
            return cellnode, sst_count
        else                                             # no more rows
            return nothing, sst_count
        end
    =#
    # unthreaded cell extraction is (exceedingly marginally) slower but no lock conflicts introduced.

    # debug
    # @assert row.tag == "row" "Not a row node"

    sst_count = 0

    d = row.depth

    cellnode = XML.next(row)

    while !isnothing(cellnode) && cellnode.depth > d
        if cellnode.tag == "c" # This is a cell
            cell = Cell(cellnode, ws; mylock) # construct an XLSX.Cell from an XML.LazyNode
            sst_count += cell.datatype == CT_STRING ? 1 : 0
            rowcells[column_number(cell)] = cell
        end
        cellnode = XML.next(cellnode)
    end
    if !isnothing(cellnode) && cellnode.tag == "row" # have reached the beginning of next row
        return cellnode, sst_count
    else                                             # no more rows
        return nothing, sst_count
    end

end
