
#
# Table
#

# Returns a tuple with the first and last index of the columns for a `SheetRow`.
function column_bounds(sr::SheetRow)

    isempty(sr) && throw(XLSXError("Can't get column bounds from an empty row."))

    local first_column_index::Int = first(keys(sr.rowcells))
    local last_column_index::Int = first_column_index

    for k in keys(sr.rowcells)
        if k < first_column_index
            first_column_index = k
        end

        if k > last_column_index
            last_column_index = k
        end
    end

    return (first_column_index, last_column_index)
end

# anchor_column will be the leftmost column of the column_bounds
function last_column_index(sr::SheetRow, anchor_column::Int) :: Int

    isempty(getcell(sr, anchor_column)) && throw(XLSXError("Can't get column bounds based on an empty anchor cell."))

    local first_column_index::Int = anchor_column
    local last_column_index::Int = first_column_index

    if length(keys(sr.rowcells)) == 1
        return anchor_column
    end

    columns = sort(collect(keys(sr.rowcells)))
    first_i = findfirst(colindex -> colindex == anchor_column, columns)
    last_column_index = anchor_column

    for i in (first_i+1):length(columns)
        if columns[i] - 1 != last_column_index
            return last_column_index
        end

        last_column_index = columns[i]
    end

    return last_column_index
end

function _colname_prefix_string(sheet::Worksheet, cell::Cell)
    d = getdata(sheet, cell)
    if d isa String
        return XML.unescape(d)
    else
        return string(d)
    end
end
_colname_prefix_string(sheet::Worksheet, ::EmptyCell) = "#Empty"

# helper function to manage problematic column labels
# Empty cell -> "#Empty"
# No_unique_label -> No_unique_label_2
function push_unique!(vect::Vector{String}, sheet::Worksheet, cell::AbstractCell, iter::Int=1)
    name = _colname_prefix_string(sheet, cell)

    if iter > 1
        name = name*"_"*string(iter)
    end

    if name in vect
        push_unique!(vect, sheet, cell, iter + 1)
    else
        push!(vect, name)
    end

    nothing
end

# Issue 260
const RESERVED = Set(["local", "global", "export", "let",
    "for", "struct", "while", "const", "continue", "import",
    "function", "if", "else", "try", "begin", "break", "catch",
    "return", "using", "baremodule", "macro", "finally",
    "module", "elseif", "end", "quote", "do"])
normalizename(name::Symbol) = name
function normalizename(name::String)::Symbol
    uname = strip(Unicode.normalize(name))
    id = Base.isidentifier(uname) ? uname : map(c->Base.is_id_char(c) ? c : '_', uname)
    cleansed = string((isempty(id) || !Base.is_id_start_char(id[1]) || id in RESERVED) ? "_" : "", id)
    return Symbol(replace(cleansed, r"(_)\1+"=>"_"))
end

"""
    eachtablerow(sheet, [columns]; [first_row], [column_labels], [header], [stop_in_empty_row], [stop_in_row_function], [keep_empty_rows], [normalizenames]) -> TableRowIterator

Constructs an iterator of table rows. Each element of the iterator is of type `TableRow`.

`header` is a boolean indicating whether the first row of the table is a table header.

If `header == false` and no `column_labels` were supplied, column names will be generated following the column names found in the Excel file.

The `columns` argument is a column range, as in `"B:E"`.
If `columns` is not supplied, the column range will be inferred by the non-empty contiguous cells in the first row of the table.

The user can replace column names by assigning the optional `column_labels` input variable with a `Vector{Symbol}`.

`stop_in_empty_row` is a boolean indicating whether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the iterator will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`. Empty rows may be returned by the iterator when `stop_in_empty_row=false`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.

Example for `stop_in_row_function`:

```
function stop_function(r)
    v = r[:col_label]
    return !ismissing(v) && v == "unwanted value"
end
```

`keep_empty_rows` determines whether rows where all column values are equal to `missing` are kept (`true`) or skipped (`false`) by the row iterator.
`keep_empty_rows` never affects the *bounds* of the iterator; the number of rows read from a sheet is only affected by `first_row`, `stop_in_empty_row` and `stop_in_row_function` (if specified).
`keep_empty_rows` is only checked once the first and last row of the table have been determined, to see whether to keep or drop empty rows between the first and the last row.

`normalizenames` controls whether column names will be "normalized" to valid Julia identifiers. By default, this is false.
If normalizenames=true, then column names with spaces, or that start with numbers, will be adjusted with underscores to become 
valid Julia identifiers. This is useful when you want to access columns via dot-access or getproperty, like file.col1. The 
identifier that comes after the . must be valid, so spaces or identifiers starting with numbers aren't allowed.
(Based ib CSV.jl's `CSV.normalizename`.)

Example code:
```
for r in XLSX.eachtablerow(sheet)
    # r is a `TableRow`. Values are read using column labels or numbers.
    rn = XLSX.row_number(r) # `TableRow` row number.
    v1 = r[1] # will read value at table column 1.
    v2 = r[:COL_LABEL2] # will read value at column labeled `:COL_LABEL2`.
end
```

See also [`XLSX.gettable`](@ref).
"""
function eachtablerow(
            sheet::Worksheet,
            cols::Union{ColumnRange, AbstractString};
            first_row::Union{Nothing, Int}=nothing,
            column_labels=nothing,
            header::Bool=true,
            stop_in_empty_row::Bool=true,
            stop_in_row_function::Union{Nothing, Function}=nothing,
            keep_empty_rows::Bool=false,
            normalizenames::Bool=false
    ) :: TableRowIterator

    if first_row === nothing
        first_row = _find_first_row_with_data(sheet, convert(ColumnRange, cols).start)
    end

    itr = eachrow(sheet)
    column_range = convert(ColumnRange, cols)
    col_lab = Vector{String}()

    if column_labels === nothing
        if header
            # will use getdata to get column names
            sheet_row = find_row(itr, first_row)
            for column_index in column_range.start:column_range.stop
                cell = getcell(sheet_row, column_index)
                push_unique!(col_lab, sheet, cell)
            end
        else
            # generate column_labels if there's no header information anywhere
            for c in column_range
                push!(col_lab, string(c))
            end
        end
    else
        # check consistency for column_range and column_labels
        if length(column_labels) != length(column_range) 
            throw(XLSXError("`column_range` (length=$(length(column_range))) and `column_labels` (length=$(length(column_labels))) must have the same length."))
        end
    end
    if normalizenames
        column_labels = normalizename.(column_labels===nothing ? col_lab : column_labels)
    else
        column_labels = Symbol.(column_labels===nothing ? col_lab : column_labels)
    end

    first_data_row = header ? first_row + 1 : first_row

    return TableRowIterator(sheet, Index(column_range, column_labels), first_data_row, stop_in_empty_row, stop_in_row_function, keep_empty_rows)
end

function TableRowIterator(sheet::Worksheet, index::Index, first_data_row::Int, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Nothing, Function}=nothing, keep_empty_rows::Bool=false)
    return TableRowIterator(eachrow(sheet), index, first_data_row, stop_in_empty_row, stop_in_row_function, keep_empty_rows)
end

function eachtablerow(
            sheet::Worksheet;
            first_row::Union{Nothing, Int}=nothing,
            column_labels=nothing,
            header::Bool=true,
            stop_in_empty_row::Bool=true,
            stop_in_row_function::Union{Nothing, Function}=nothing,
            keep_empty_rows::Bool=false,
            normalizenames::Bool=false
        ) :: TableRowIterator

        if first_row === nothing
        # if no columns were given,
        # first_row must be provided and cannot be inferred.
        # If it was not provided, will use first row as default value
        first_row = 1
    end

    for r in eachrow(sheet)

        # skip rows until we reach first_row, and if !keep_empty_rows then skip empty rows
        if row_number(r) < first_row || isempty(r) && !keep_empty_rows
            continue
        end

        columns_ordered = sort(collect(keys(r.rowcells)))

        for (ci, cn) in enumerate(columns_ordered)
            if !ismissing(getdata(r, cn))
                # found a row with data. Will get ColumnRange from non-empty consecutive cells
                first_row = row_number(r)
                column_start = cn
                column_stop = cn

                if length(columns_ordered) == 1
                    # there's only one column
                    column_range = ColumnRange(column_start, column_stop)
                    return eachtablerow(sheet, column_range; first_row, column_labels, header=header, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames)
                else
                    # will figure out the column range
                    for ci_stop in (ci+1):length(columns_ordered)
                        cn_stop = columns_ordered[ci_stop]

                        # Will stop if finds an empty cell or a skipped column
                        if ismissing(getdata(r, cn_stop)) || (cn_stop - 1 != column_stop)
                            column_range = ColumnRange(column_start, column_stop)
                            return eachtablerow(sheet, column_range; first_row, column_labels, header, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames)
                        end
                        column_stop = cn_stop
                    end
                end

                # if got here, it's because all columns are non-empty
                column_range = ColumnRange(column_start, column_stop)
                return eachtablerow(sheet, column_range; first_row, column_labels, header, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames)
            end
        end
    end

    throw(XLSXError("Couldn't find a table in sheet $(sheet.name)"))
end

function _find_first_row_with_data(sheet::Worksheet, column_number::Int)
    # will find first_row
    for r in eachrow(sheet)
        if !ismissing(getdata(r, column_number))
            return row_number(r)
        end
    end
    throw(XLSXError("Column $(encode_column_number(column_number)) has no data."))
end

@inline get_worksheet(tri::TableRowIterator) = get_worksheet(tri.itr)

# Returns real sheet column numbers (based on cellref)
@inline sheet_column_numbers(i::Index) = values(i.column_map)

# Returns an iterator for table column numbers.
@inline table_column_numbers(i::Index) = eachindex(i.column_labels)
@inline table_column_numbers(r::TableRow) = table_column_numbers(r.index)

# Maps table column index (1-based) -> sheet column index (cellref based)
@inline table_column_to_sheet_column_number(index::Index, table_column_number::Int) = index.column_map[table_column_number]
@inline table_columns_count(i::Index) = length(i.column_labels)
@inline table_columns_count(itr::TableRowIterator) = table_columns_count(itr.index)
@inline table_columns_count(r::TableRow) = table_columns_count(r.index)
@inline row_number(r::TableRow) = r.row
@inline get_column_labels(index::Index) = index.column_labels
@inline get_column_labels(itr::TableRowIterator) = get_column_labels(itr.index)
@inline get_column_labels(r::TableRow) = get_column_labels(r.index)
@inline get_column_label(r::TableRow, table_column_number::Int) = get_column_labels(r)[table_column_number]

# iterate into TableRow to get each column value

function Base.iterate(r::TableRow)
    next = iterate(table_column_numbers(r))
    if next === nothing
        return nothing
    else
        next_column_number, next_state = next
        return r[next_column_number], next_state
    end
end

function Base.iterate(r::TableRow, state)
    next = iterate(table_column_numbers(r), state)
    if next === nothing
        return nothing
    else
        next_column_number, next_state = next
        return r[next_column_number], next_state
    end
end

Base.getindex(r::TableRow, x) = getdata(r, x)

function TableRow(table_row::Int, index::Index, sheet_row::SheetRow)
    ws = get_worksheet(sheet_row)

    cell_values = Vector{CellValueType}()
    for table_column_number in table_column_numbers(index)
        sheet_column = table_column_to_sheet_column_number(index, table_column_number)
        cell = getcell(sheet_row, sheet_column)
        push!(cell_values, getdata(ws, cell))
    end

    return TableRow(table_row, index, cell_values)
end

getdata(r::TableRow, table_column_number::Int) = r.cell_values[table_column_number]
getdata(r::TableRow, table_column_numbers::Union{Vector{T}, UnitRange{T}}) where {T<:Integer} = [r.cell_values[x] for x in table_column_numbers]

function getdata(r::TableRow, column_label::Symbol)
    index = r.index
    if haskey(index.lookup, column_label)
        return getdata(r, index.lookup[column_label])
    else
        throw(XLSXError("Invalid column label: $column_label."))
    end
end

Base.IteratorSize(::Type{<:TableRowIterator}) = Base.SizeUnknown()
Base.eltype(::TableRowIterator) = TableRow

function Base.iterate(itr::TableRowIterator)
    next = iterate(itr.itr)

    # go to the first_data_row
    while next !== nothing
        (sheet_row, sheet_row_iterator_state) = next

        if row_number(sheet_row) == itr.first_data_row
            table_row_index = 1
            missing_rows=0
            return TableRow(table_row_index, itr.index, sheet_row), TableRowIteratorState(table_row_index, row_number(sheet_row), sheet_row_iterator_state, missing_rows, nothing)
        else
           next = iterate(itr.itr, sheet_row_iterator_state)
        end
    end

    # no rows for this table
    return nothing
end

function Base.iterate(itr::TableRowIterator, state::TableRowIteratorState)
    table_row_index = state.table_row_index + 1
    missing_rows=state.missing_rows
    col_count=length(sheet_column_numbers(itr.index))

    if missing_rows > 0 # sheetrow iterator has skipped some completely empty rows
        if itr.stop_in_empty_row
            # user asked to stop fetching table rows if we find an empty row
            println("Shouldn't see this message") # handled below
            return nothing
        elseif itr.keep_empty_rows
            # return a TableRow with missing values for the columns
            table_row = TableRow(table_row_index, itr.index, fill(missing, col_count))
            table_row_index += 1
            missing_rows -= 1
            return table_row, TableRowIteratorState(table_row_index, state.sheet_row_index, state.sheet_row_iterator_state, missing_rows, state.row_pending)
        else
            throw(XLSXError("Something wrong here"))
        end
    elseif isnothing(state.row_pending)
        # Only interate sheetrow if we've properly handled any entirely empty rows.
        next = iterate(itr.itr, state.sheet_row_iterator_state) # iterate the SheetRowIterator
        if next === nothing
            return nothing
        end
        sheet_row, sheet_row_iterator_state = next
    else
        # bring forward the pending row
        sheet_row_iterator_state = state.sheet_row_iterator_state
        sheet_row = state.row_pending
    end

    #
    # checks if we're done reading this table
    #

    # check skipping rows
    # The XML can skip rows if there's no data in it,
    # so this is why is_empty_table_row function below wouldn't catch this case
    if itr.stop_in_empty_row && row_number(sheet_row) != itr.first_data_row && row_number(sheet_row) != (state.sheet_row_index + 1)
        return nothing
    end

    # checks if there are any data inside column range (row not entirely empty)
    function is_empty_table_row(itr::TableRowIterator, sheet_row::SheetRow) :: Bool
        if isempty(sheet_row)
            return true
        end

        for c in sheet_column_numbers(itr.index)
            if !ismissing(getdata(get_worksheet(itr), getcell(sheet_row, c)))
                return false
            end
        end
        return true
    end

    if !isnothing(state.row_pending)
        # bring forward pending row
        table_row = TableRow(table_row_index, itr.index, state.row_pending)
        newstate = TableRowIteratorState(table_row_index, row_number(sheet_row), sheet_row_iterator_state, 0, nothing)
    elseif !itr.stop_in_empty_row && itr.keep_empty_rows && row_number(sheet_row) != (state.sheet_row_index + 1)
        # the sheetrow iterator has skipped some empty rows. Postpone processing this sheet row and process empty rows if keep_empty_rows is true
        missing_rows = row_number(sheet_row) - state.sheet_row_index - 1
        table_row = TableRow(table_row_index, itr.index, fill(missing, col_count))
        missing_rows -= 1
        row_pending = sheet_row
        newstate = TableRowIteratorState(table_row_index, state.sheet_row_index, state.sheet_row_iterator_state, missing_rows, row_pending)
    else
        # normal case, no empty rows
        table_row = TableRow(table_row_index, itr.index, sheet_row)
        newstate = TableRowIteratorState(table_row_index, row_number(sheet_row), sheet_row_iterator_state, 0, nothing)
    end

    if is_empty_table_row(itr, sheet_row) # rows are returned but specified columns are empty
        if itr.stop_in_empty_row 
            # user asked to stop fetching table rows if we find an empty row
            return nothing
        elseif !itr.keep_empty_rows
            # keep looking for a non-empty row
            next = iterate(itr.itr, sheet_row_iterator_state)
            while next !== nothing
                sheet_row, sheet_row_iterator_state = next
                if !is_empty_table_row(itr, sheet_row)
                    break
                end
                next = iterate(itr.itr, sheet_row_iterator_state)
            end

            if next === nothing
                # end of file
                return nothing
            end
            table_row = TableRow(table_row_index, itr.index, sheet_row)
            newstate = TableRowIteratorState(table_row_index, row_number(sheet_row), sheet_row_iterator_state, 0, nothing)
        end
    end

    if itr.stop_in_row_function !== nothing && itr.stop_in_row_function(table_row)
        return nothing
    end

    return table_row, newstate

end

function infer_eltype(v::Vector{Any})
    local hasmissing::Bool = false
    local t::DataType = Any

    if isempty(v)
        return eltype(v)
    end

    for i in 1:length(v)
        if ismissing(v[i])
            hasmissing = true
        else
            if t != Any && typeof(v[i]) != t
                t = promote_type(t, typeof(v[i]))
                if t == Any
                    return t
                end
                # return Any
            else
                t = typeof(v[i])
            end
        end
    end

    if t == Any
        return Any
    else
        if hasmissing
            return Union{Missing, t}
        else
            return t
        end
    end
end

infer_eltype(v::Vector{T}) where T = T

function check_table_data_dimension(data::Vector)

    # nothing to check
    isempty(data) && return

    # all columns should be vectors
    for (colindex, colvec) in enumerate(data)
        if !isa(colvec, Vector)
            throw(XLSXError("Data type at index $colindex is not a vector. Found: $(typeof(colvec))."))
        end
    end

    # no need to check row count
    length(data) == 1 && return

    # check all columns have the same row count
    col_count = length(data)
    row_count = length(data[1])
    for colindex in 2:col_count
        if length(data[colindex]) != row_count
            throw(XLSXError("Not all columns have the same number of rows. Check column $colindex."))
        end
    end

    nothing
end

function gettable(itr::TableRowIterator; infer_eltypes::Bool=true) :: DataTable

    column_labels = get_column_labels(itr)
    columns_count = table_columns_count(itr)
    data = Vector{Any}(undef, columns_count)
    for c in 1:columns_count
        data[c] = Vector{Any}()
    end

    for r in itr # r is a TableRow
        is_empty_row = true
        for (ci, cv) in enumerate(r) # iterate a TableRow to get column data
            push!(data[ci], cv)
            if !ismissing(cv)
                is_empty_row = false
            end
        end

        # undo insert row in case of empty rows
        if is_empty_row && (itr.stop_in_empty_row || !itr.keep_empty_rows)
            for c in 1:columns_count
                pop!(data[c])
            end
        end
    end

    if infer_eltypes
        rows = length(data[1])
        for c in 1:columns_count
            new_column_data = Vector{infer_eltype(data[c])}(undef, rows)
            for r in 1:rows
                new_column_data[r] = data[c][r]
            end
            data[c] = new_column_data
        end
    end

    check_table_data_dimension(data)

    return DataTable(data, column_labels)
end

"""
    gettable(
        sheet,
        [columns];
        [first_row],
        [column_labels],
        [header],
        [infer_eltypes],
        [stop_in_empty_row],
        [stop_in_row_function],
        [keep_empty_rows],
        [normalizenames]
    ) -> DataTable

Returns data from a spreadsheet as a struct `XLSX.DataTable` which
can be passed directly to any function that accepts `Tables.jl` data.
(e.g. `DataFrame` from package `DataFrames.jl`).

Use `columns` argument to specify which columns to get.
For example, `"B:D"` will select columns `B`, `C` and `D`.
If `columns` is not given, the algorithm will find the first sequence
of consecutive non-empty cells.

Use `first_row` to indicate the first row from the table.
`first_row=5` will look for a table starting at sheet row `5`.
If `first_row` is not given, the algorithm will look for the first
non-empty row in the spreadsheet.

`header` is a `Bool` indicating if the first row is a header.
If `header=true` and `column_labels` is not specified, the column labels
for the table will be read from the first row of the table.
If `header=false` and `column_labels` is not specified, the algorithm
will generate column labels. The default value is `header=true`.

Use `column_labels` as a vector of symbols to specify names for the header of the table.

Use `normalizenames=true` to normalize column names to valid Julia identifiers.

Use `infer_eltypes=true` to get `data` as a `Vector{Any}` of typed vectors.
The default value is `infer_eltypes=true`.

`stop_in_empty_row` is a boolean indicating whether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the `TableRowIterator` will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.

# Example for `stop_in_row_function`

```julia
function stop_function(r)
    v = r[:col_label]
    return !ismissing(v) && v == "unwanted value"
end
```

`keep_empty_rows` determines whether rows where all column values are equal to `missing` are kept (`true`) or dropped (`false`) from the resulting table.
`keep_empty_rows` never affects the *bounds* of the table; the number of rows read from a sheet is only affected by `first_row`, `stop_in_empty_row` and `stop_in_row_function` (if specified).
`keep_empty_rows` is only checked once the first and last row of the table have been determined, to see whether to keep or drop empty rows between the first and the last row.

# Example

```julia
julia> using DataFrames, PrettyTables, XLSX

julia> df = XLSX.openxlsx("myfile.xlsx") do xf
        DataFrame(XLSX.gettable(xf["mysheet"]))
    end

julia> PrettyTable(XLSX.gettable(xf["mysheet"], "A:C"))
┌─────────┬─────────┬─────────┐
│ Header1 │ Header2 │ Header3 │
├─────────┼─────────┼─────────┤
│       1 │       2 │       3 │
│       4 │       5 │       6 │
│       7 │       8 │       9 │
└─────────┴─────────┴─────────┘
   
```

See also: [`XLSX.readtable`](@ref), [`XLSX.readto`](@ref).
"""
function gettable(sheet::Worksheet, cols::Union{ColumnRange, AbstractString}; first_row::Union{Nothing, Int}=nothing, column_labels=nothing, header::Bool=true, infer_eltypes::Bool=true, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Nothing}=nothing, keep_empty_rows::Bool=false, normalizenames::Bool=false)
    itr = eachtablerow(sheet, cols; first_row, column_labels, header, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames)
    return gettable(itr; infer_eltypes)
end

function gettable(sheet::Worksheet; first_row::Union{Nothing, Int}=nothing, column_labels=nothing, header::Bool=true, infer_eltypes::Bool=true, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Nothing}=nothing, keep_empty_rows::Bool=false, normalizenames::Bool=false)
    itr = eachtablerow(sheet; first_row, column_labels, header, stop_in_empty_row, stop_in_row_function, keep_empty_rows, normalizenames)
    return gettable(itr; infer_eltypes)
end

#---------------------------------------------------------------------------------------------------------------------------- Transposed Tables

function transposetable(m::Matrix; header::Bool=true) # transpose a matrix and extract to vector of vectors (data columns) and vector of column names
    
    v = collect(PermutedDimsArray(m, (2, 1))) # transpose rows and columns

    #identify leading and trailing missing rows
    stop_row = axes(v, 1)[end]
    start_row = 0
    for i in 1:stop_row
        if all(ismissing, v[i, :])
            if start_row != 0
                stop_row = i-1
                break
            end
        else
            start_row==0 && (start_row=i)
        end
    end

    #identify leading and trailing missing columns
    stop_col = axes(v, 2)[end]
    start_col = 0
    for i in 1:stop_col
        if all(ismissing, v[:, i])
            if start_col != 0
                stop_col = i-1
                break
            end
        else
            start_col==0 && (start_col=i)
        end
    end

    if header # separate the header row if present
        cols = v[start_row+1:stop_row, start_col:stop_col]
        headers = (v[start_row, start_col:stop_col])
    else
        cols = v[start_row:stop_row, start_col:stop_col]
        headers = []
    end
    data = [] # convert matrix to vector of vectors and infer types
    for c in axes(cols, 2) 
        T = infer_eltype(cols[:, c])
        if T !== Any
            d = convert(Vector{T}, cols[:, c])
        else
            d = cols[:, c]
        end
        push!(data, d)                                                                                                                                                                      
    end
    return data, headers
end

"""
    gettransposedtable(
        sheet,
        [rows];
        [first_column],
        [column_labels],
        [header],
        [normalizenames]
    ) -> DataTable

Read a transposed table from a worksheet in which data are arranged in 
rows rather than columns. For example:
```
Category    "A", "B", "C", "D"
variable 1  10,  20,  30,  40
variable 2  15,  25,  35,  40
variable 3  20,  30,  40,  50
```
Returns data from a worksheet as a struct `XLSX.DataTable` which
can be passed directly to any function that accepts `Tables.jl` data.
(e.g. `DataFrame` from package `DataFrames.jl`).

Use the `rows` argument to specify which worksheeet rows to include.
For example, `"2:7"` will select rows 2 to 7 (inclusive).
If `rows` is not given, the algorithm will find the first sequence
of consecutive non-empty cells. If `rows` includes leading or trailing 
rows that are completely empty, these rows will be omitted from the 
returned table. In any case, the table will be truncated at the first 
and last non-empty rows, even if this range is smaller than `rows`. 
A valid `sheet` must be specified when specifying `rows`.

Use `first_column` to indicate the first column of the table. May be given 
as a column number or as a string, so that `first_column="E"` and
`first_column=5` will both look for a table starting at column `5` ("E").
Any leading completely empty columns will be ignored, including 
the `first_column`. If `first_column` is not given, the algorithm will 
look for the first non-empty column in the spreadsheet.

`header` is a `Bool` indicating if the first row is a header.
If `header=true` and `column_labels` is not specified, the column labels
for the table will be read from the first column of the table.
If `header=false` and `column_labels` is not specified, the algorithm
will generate column labels. The default value is `header=true`.

Use `column_labels` as a vector of symbols to specify names for the 
header of the table. If `header=true` and `column_labels` is also given, 
column_labels will be preferred and the first column of the table will 
be ignored.

Use `normalizenames=true` to normalize column names to valid Julia identifiers. 
The default is `normalizenames=false`

# Examples

```julia
julia> using DataFrames, PrettyTables, XLSX

julia> xf = XLSX.openxlsx("HTable.xlsx")
XLSXFile("HTable.xlsx") containing 4 Worksheets
            sheetname size          range
-------------------------------------------------
               Origin 6x10          B2:K7
               Offset 8x12          A1:L8
             Multiple 8x22          A1:V8
              Example 4x5           B2:F5
              
julia> DataFrame(XLSX.gettransposedtable(xf["Example"]))
4×4 DataFrame
 Row │ Category  Variable 1  Variable 2  Variable 3 
     │ String    Int64       Int64       Int64
─────┼──────────────────────────────────────────────
   1 │ A                 10          15          20
   2 │ B                 20          25          30
   3 │ C                 30          35          40
   4 │ D                 40          40          50

julia> PrettyTable(XLSX.gettransposedtable(xf["Example"]; normalizenames=true))
┌──────────┬────────────┬────────────┬────────────┐
│ Category │ Variable_1 │ Variable_2 │ Variable_3 │
├──────────┼────────────┼────────────┼────────────┤
│        A │         10 │         15 │         20 │
│        B │         20 │         25 │         30 │
│        C │         30 │         35 │         40 │
│        D │         40 │         40 │         50 │
└──────────┴────────────┴────────────┴────────────┘

julia> DataFrame(gettransposedtable(xf["Example"]; header=false))
5×4 DataFrame
 Row │ Col_1     Col_2       Col_3       Col_4      
     │ String    Any         Any         Any
─────┼──────────────────────────────────────────────
   1 │ Category  Variable 1  Variable 2  Variable 3
   2 │ A         10          15          20
   3 │ B         20          25          30
   4 │ C         30          35          40
   5 │ D         40          40          50

```
The worksheet `Multiple` contains two tables side by side, separated by an empty column.
Only the first table is read by default. Read the second table by additionally specifying 
the `first_column`.

```julia
julia> DataFrame(XLSX.gettransposedtable(xf["Multiple"], "2:7"))
9×6 DataFrame
 Row │ Year   Col A  Col B  Col C  Col D    Col E      
     │ Int64  Int64  Int64  Int64  Float64  Any
─────┼─────────────────────────────────────────────────
   1 │  1940      1     10    100      0.1  Hello
   2 │  1950      2     20    200      0.2  2025-12-19
   3 │  1960      3     30    300      0.3  3
   4 │  1970      4     40    400      0.4  3.33
   5 │  1980      5     50    500      0.5  Hello
   6 │  1990      6     60    600      0.6  2025-12-19
   7 │  2000      7     70    700      0.7  3
   8 │  2010      8     80    800      0.8  3.33
   9 │  2020      9     90    900      0.9  true

julia> DataFrame(XLSX.gettransposedtable(xf["Multiple"], "2:7"; first_column="M"))
9×6 DataFrame
 Row │ date   name1    name2    name3  name4     name5      
     │ Int64  Float64  Float64  Bool   Time      Any
─────┼──────────────────────────────────────────────────────
   1 │  1840     12.4    0.045   true  10:22:00  Hello
   2 │  1841     12.6    0.046   true  10:23:00  2025-12-19
   3 │  1842     12.8    0.047  false  10:24:00  3
   4 │  1843     13.0    0.048   true  10:25:00  3.33
   5 │  1844     13.2    0.049  false  10:26:00  Hello
   6 │  1845     13.4    0.05    true  10:27:00  2025-12-19
   7 │  1846     13.6    0.051   true  10:28:00  3
   8 │  1847     13.8    0.052   true  10:29:00  3.33
   9 │  1848     14.0    0.053  false  10:30:00  true

```

See also: [`XLSX.readtransposedtable`](@ref), [`XLSX.readtable`](@ref).
"""
function gettransposedtable(sheet::Worksheet, rows::Union{AbstractString,Nothing}=nothing; first_column=nothing, column_labels=nothing, header::Bool=true, normalizenames::Bool=false)
    dim = get_dimension(sheet)
    if isnothing(rows)
        rng=RowRange(dim.start.row_number, dim.stop.row_number) 
    else
        is_valid_row_range(rows) || throw(XLSXError("Invalid row range: $rows"))
        rng=RowRange(rows)
    end
    if rng.start < dim.start.row_number || rng.stop > dim.stop.row_number
        throw(XLSXError("Row range $rows extends outside sheet dimension ($(dim.start.row_number):$(dim.stop.row_number))"))
    end
    if first_column isa String
        first_column=decode_column_number(first_column)
    elseif first_column !== nothing && !(first_column isa Int)
        throw(XLSXError("first_column must be an integer column number or a column string like \"A\", \"B\", etc."))
    end
    if isnothing(first_column)
        first_column=0
    else
        if (first_column > dim.stop.column_number || first_column < dim.start.column_number)
            throw(XLSXError("First column $first_column ($(encode_column_number(first_column))) is outside of sheet dimension ($(dim.start.column_number):$(dim.stop.column_number))"))
        end
    end
    start = CellRef(rng.start, max(dim.start.column_number, first_column))
    stop = CellRef(rng.stop, dim.stop.column_number)
    m = sheet[CellRange(start, stop)]
    data, h = transposetable(m; header)
    if isnothing(column_labels)
        if header==true
            column_labels=h
        else
            column_labels=["Col_$(i)" for i in 1:length(data)]
        end
    end
    if normalizenames
        column_labels = Symbol.(normalizename.(column_labels))
    else
        column_labels = Symbol.(column_labels)
    end
    check_table_data_dimension(data)

    return DataTable(data, column_labels)
end

"""
    XLSXFile(table)

Take a `Tables.jl` compatible table and create a new `XLSXFile` object for writing.
Can act as a sink for functions such as `CSV.read`.

# Example

```julia
julia> using CSV, XLSX

julia> xf = CSV.read("iris.csv", XLSXFile)
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 151x5         A1:E151
```

"""
function XLSXFile(table)
    xf=newxlsx()
    writetable!(xf[1], table)
    return xf
end