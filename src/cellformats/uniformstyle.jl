#
# -- Set uniform styles
#

"""
    setUniformStyle(sh::Worksheet, cr::String) -> ::Int
    setUniformStyle(xf::XLSXFile,  cr::String) -> ::Int

    setUniformStyle(sh::Worksheet, rows, cols) -> ::Int

Set the cell `style` used by a cell range, a column range or row range 
or a named range in a worksheet or XLSXFile to be the same as that of 
the first cell in the range that is not an `EmptyCell`.
Alternatively, specify the rows and columns using any combination of 
Integer, UnitRange, Vector{Integer} or `:`.

As a result, every cell in the range will have a uniform `style`.

A cell `style` consists of the collection of `format`, `alignment`, `border`, 
`font` and `fill`.

If the first cell has no defined `style` (`s=""`), all cells will be given the 
same undefined `style`.

The value returned is the `styleId` of the `style` uniformly applied to the cells or 
`nothing` if the style is undefined.
If all cells in the range are `EmptyCells`, the returned value is -1.

# Examples:
```julia
julia> XLSX.setUniformStyle(xf, "Sheet1!A2:L6")

julia> XLSX.setUniformStyle(sh, "F1:F5")

julia> XLSX.setUniformStyle(sh, 2:5, 5)

julia> XLSX.setUniformStyle(sh, 2, :)
 
```
"""
function setUniformStyle end
setUniformStyle(ws::Worksheet, rng::SheetCellRange) = do_sheet_names_match(ws, rng) && setUniformStyle(ws, rng.rng)
setUniformStyle(ws::Worksheet, rng::SheetColumnRange) = do_sheet_names_match(ws, rng) && setUniformStyle(ws, rng.colrng)
setUniformStyle(ws::Worksheet, rng::SheetRowRange) = do_sheet_names_match(ws, rng) && setUniformStyle(ws, rng.rowrng)
setUniformStyle(ws::Worksheet, colrng::ColumnRange) = process_columnranges(setUniformStyle, ws, colrng)
setUniformStyle(ws::Worksheet, rowrng::RowRange) = process_rowranges(setUniformStyle, ws, rowrng)
setUniformStyle(ws::Worksheet, ncrng::NonContiguousRange) = process_uniform_ncranges(ws, ncrng)
setUniformStyle(xl::XLSXFile, sheetcell::AbstractString) = process_sheetcell(setUniformStyle, xl, sheetcell)
setUniformStyle(ws::Worksheet, ref_or_rng::AbstractString) = process_ranges(setUniformStyle, ws, ref_or_rng)
setUniformStyle(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon) = process_colon(ws, row, nothing)
setUniformStyle(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}) = process_colon(ws, nothing, col)
setUniformStyle(ws::Worksheet, ::Colon, ::Colon) = process_colon(ws, nothing, nothing)
setUniformStyle(ws::Worksheet, ::Colon) = process_colon(ws, nothing, nothing)
setUniformStyle(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon) = process_uniform_veccolon(ws, row, nothing)
setUniformStyle(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}) = process_uniform_veccolon(ws, nothing, col)
setUniformStyle(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = process_uniform_vecint(ws, row, col)
setUniformStyle(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = process_uniform_vecint(ws, row, col)
setUniformStyle(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}) = process_uniform_vecint(ws, row, col)
setUniformStyle(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = setUniformStyle(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))

function setUniformStyle(ws::Worksheet, rng::CellRange)::Union{Nothing,Int}
    _require_writable(ws, "set styles")

    newid::Union{Nothing,Int} = nothing
    is_first   = true
    first_font = nothing

    for cellref in rng
        newid, is_first, first_font = process_uniform_core(ws, cellref, newid, is_first, first_font)
    end

    # -1 signals that all cells in the range were EmptyCells
    is_first && return -1

    return newid
end
