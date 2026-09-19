
struct CellPosition
    row::Int
    column::Int
end

"""
    CellRef(n::AbstractString)
    CellRef(row::Int, col::Int)

A `CellRef` represents a cell location given by row and column identifiers.

`CellRef("B6")` indicates a cell located at column `2` and row `6`.

These row and column integers can also be passed directly to the `CellRef` constructor: `CellRef(6,2) == CellRef("B6")`.

Finally, a convenience macro `@ref_str` is provided: `ref"B6" == CellRef("B6")`.

# Examples

```julia
cn = XLSX.CellRef("AB1")
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
println( cellname(cn) ) # will print out AB1

cn = XLSX.CellRef(1, 28)
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
println( cellname(cn) ) # will print out AB1

cn = XLSX.ref"AB1"
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
println( cellname(cn) ) # will print out AB1
```

"""
struct CellRef
    row_number::Int32
    column_number::Int32
end

abstract type AbstractCellDataFormat end

struct EmptyCellDataFormat <: AbstractCellDataFormat end

# Keeps track of formatting information.
struct CellDataFormat <: AbstractCellDataFormat
    id::UInt32
end

abstract type AbstractFormula end
abstract type ExplicitFormula <: AbstractFormula end

struct EmptyFormula <: AbstractFormula end

"""
A default formula simply storing the formula string.
"""
mutable struct Formula <: ExplicitFormula
    formula::String
    type::Union{String,Nothing} # usually nothing but has value "array" for dynamic array functions.
    ref::Union{String,Nothing} # usually nothing but refers to the "spill" range for dynamic array functions.
    unhandled::Union{Dict{String,String},Nothing}
end
function Formula()
    return EmptyFormula()
end
function Formula(s::String)
    return Formula(s, nothing, nothing, nothing)
end


"""
The formula in this cell was defined somewhere else; we simply reference its ID.
"""
mutable struct FormulaReference <: AbstractFormula
    id::Int
    unhandled::Union{Dict{String,String},Nothing}
end

"""
Formula that is defined once and referenced in all cells given by the cell range given in `ref` and using the same `id`.
"""
mutable struct ReferencedFormula <: ExplicitFormula
    formula::String
    id::Int
    ref::String # actually a CellRange, but defined later --> change if at some point we want to actively change formulae
    unhandled::Union{Dict{String,String},Nothing}
end

struct CellFormula# <: AbstractFormula
    value::T where T<:AbstractFormula
    styleid::AbstractCellDataFormat
end

# Keeps track of external references in formulas.
struct ExternalRef
    index::Int          # the [n] index in the formula
    sheet::String       # sheet name
    full::String        # raw "[n]Sheet!$A$1" formula element
end


mutable struct CellFont
    fontId::Int
    font::Dict{String, Union{Dict{String, String}, Nothing}} # fontAttribute -> (attribute -> value)
    applyFont::String

    function CellFont(fontid::Int, font::Dict{String, Union{Dict{String, String}, Nothing}}, applyFont::String)
        return new(fontid, font, applyFont)
    end
end

# A border postion element (e.g. `top` or `left`) has a style attribute, but `color` is a child element.
# The `color` element has an attribute (e.g. `rgb`) that defines the color of the border.
# These are both stored in the `border` field of `CellBorder`. The key for the color element
# will vary depending on how the color is defined (e.g. `rgb`, `indexed`, `auto`, etc.).
# Thus, for example, `"top" => Dict("style" => "thin", "rgb" => "FF000000")`
mutable struct CellBorder
    borderId::Int
    border::Dict{String, Union{Dict{String, String}, Nothing}} # borderAttribute -> (attribute -> value)
    applyBorder::String

    function CellBorder(borderid::Int, border::Dict{String, Union{Dict{String, String}, Nothing}}, applyBorder::String)
        return new(borderid, border, applyBorder)
    end
end

# A fill has a pattern type attribute and two children fgColor and bgColor, each with 
# one or two attributes of their own. These color attributes are pushed in to the Dict 
# of attributes with either `fg` or `bg` prepended to their name to support later 
# reconstruction of the xml element.
mutable struct CellFill
    fillId::Int
    fill::Dict{String, Union{Dict{String, String}, Nothing}} # fillAttribute -> (attribute -> value)
    applyFill::String

    function CellFill(fillid::Int, fill::Dict{String, Union{Dict{String, String}, Nothing}}, applyfill::String)
        return new(fillid, fill, applyfill)
    end
end
mutable struct CellFormat
    numFmtId::Int
    format::Dict{String, Union{Dict{String, String}, Nothing}} # fillAttribute -> (attribute -> value)
    applyNumberFormat::String

    function CellFormat(formatid::Int, format::Dict{String, Union{Dict{String, String}, Nothing}}, applynumberformat::String)
        return new(formatid, format, applynumberformat)
    end
end

mutable struct CellAlignment # Alignment is part of the cell style `xf` so doesn't need an Id
    alignment::Dict{String, Union{Dict{String, String}, Nothing}} # alignmentAttribute -> (attribute -> value)
    applyAlignment::String

    function CellAlignment(alignment::Dict{String, Union{Dict{String, String}, Nothing}}, applyalignment::String)
        return new(alignment, applyalignment)
    end
end

@enum CellValueType::UInt8 begin
    CT_EMPTY = 0
    CT_STRING = 1
    CT_FLOAT = 2
    CT_INT = 3
    CT_BOOL = 4
    CT_DATE = 5
    CT_TIME = 6
    CT_DATETIME = 7
    CT_ERROR = 8
end
@enum CellErrorType::UInt64 begin
    XL_NULL = 1
    XL_DIV0 = 2
    XL_VALUE = 3
    XL_REF = 4
    XL_NAME = 5
    XL_NUM = 6 
    XL_NA = 7
    XL_SPILL = 8 # Turns out #SPILL is not an official error. These will return #VALUE errors
end

abstract type AbstractCell end

mutable struct Cell <: AbstractCell
    ref::CellRef 
    value::UInt64 # Needs to be `reinterpret`ed according to the `Cell.datatype`
    style::UInt32
    meta::UInt16
    datatype::CellValueType # ENUM determines how `Cell.value` needs to be `reinterpret`ed
    formula::Bool # has a formula in Workbook.formulas
end

struct EmptyCell <: AbstractCell
    ref::CellRef
end

# Keeps track of conditional formatting information.
struct DxFormat <: AbstractCellDataFormat
    id::UInt
end

"""
    CellConcreteType

Concrete supported data-types.

```julia
Union{String, Missing, Float64, Int64, Bool, Dates.Date, Dates.Time, Dates.DateTime}
```

!!! note

    In julia, the values `Inf`, `-Inf` and `NaN` are of type `Float64`. However, there is 
    no way to represent these values as numbers in Excel. Instead, these specific values 
    are eagerly converted to string representation (`"Inf"`, `"-Inf"` and `"NaN"`) as they 
    are added to an XLSXFile and they are written out as such to any saved `.xlsx` file.

"""
const CellConcreteType = Union{String, Missing, Float64, Int64, Bool, Dates.Date, Dates.Time, Dates.DateTime}
# `Int64`, not `Int`: the set of types a cell can hold must not depend on the
# host word size, or 32-bit builds would silently narrow it to `Int32`.
# `setdata!` normalises every `Integer` to `Int64`, so consumers testing a cell
# value's type must test against `Int64` too (see write.jl).

# CellValue is a Julia type of a value read from a Spreadsheet.
struct CellValue
    value::CellConcreteType
    styleid::AbstractCellDataFormat
    CellValue(value::CellConcreteType, styleid::AbstractCellDataFormat) = new(value, styleid)
    CellValue(value::Integer, styleid::AbstractCellDataFormat) = new(Int64(value), styleid)
end

#=
A `CellRange` represents a rectangular range of cells in a spreadsheet.

`CellRange("A1:C4")` denotes cells ranging from `A1` (upper left corner) to `C4` (bottom right corner).

As a convenience, `@range_str` macro is provided.

```julia
cr = XLSX.range"A1:C4"
```
=#

abstract type AbstractCellRange end
abstract type ContiguousCellRange <: AbstractCellRange end
abstract type AbstractSheetCellRange <: AbstractCellRange end
abstract type ContiguousSheetCellRange <: AbstractSheetCellRange end

struct CellRange <: ContiguousCellRange
    start::CellRef
    stop::CellRef

    function CellRange(a::CellRef, b::CellRef)

        top = row_number(a)
        bottom = row_number(b)
        left = column_number(a)
        right = column_number(b)

        if left > right || top > bottom
            throw(XLSXError("Invalid CellRange. Start cell should be at the top left corner of the range."))
        end

        return new(a, b)
    end
end

struct ColumnRange <: ContiguousCellRange
    start::Int # column number
    stop::Int  # column number

    function ColumnRange(a::Integer, b::Integer)
        if a > b 
            throw(XLSXError("Invalid ColumnRange. Start column must be located before end column."))
        end
        return new(a, b)
    end
end
struct RowRange <: ContiguousCellRange
    start::Int # row number
    stop::Int  # row number

    function RowRange(a::Integer, b::Integer)
        if a > b
            throw(XLSXError("Invalid RowRange. Start row must be located before end row."))
        end
        return new(a, b)
    end
end

struct SheetCellRef
    sheet::String
    cellref::CellRef
end

struct SheetCellRange <: ContiguousSheetCellRange
   sheet::String
   rng::CellRange
end

struct NonContiguousRange <: AbstractSheetCellRange
    sheet::String
    rng::Vector{Union{CellRef, CellRange}}
end

struct SheetColumnRange <: ContiguousSheetCellRange
    sheet::String
    colrng::ColumnRange
end
struct SheetRowRange <: ContiguousSheetCellRange
    sheet::String
    rowrng::RowRange
end

abstract type MSOfficePackage end

struct EmptyMSOfficePackage <: MSOfficePackage
end

#=
Relationships are defined in ECMA-376-1 Section 9.2.
This struct matches the `Relationship` tag attribute names.

A `Relationship` defines relations between the files inside a MSOffice package.
Regarding Spreadsheets, there are two kinds of relationships:

    * package level: defined in `_rels/.rels`.
    * workbook level: defined in `xl/_rels/workbook.xml.rels`.

The function `parse_relationships!(xf::XLSXFile)` is used to parse
package and workbook level relationships.
=#
struct Relationship
    Id::String
    Type::String
    Target::String
end

const CellCache = Dict{Int, Dict{Int, Cell}} # row -> ( column -> cell )

#=
Iterates over Worksheet cells. See `eachrow` method docs.
Each element is a `SheetRow`.

Implementations: SheetRowStreamIterator, WorksheetCache.
=#
abstract type SheetRowIterator end

mutable struct SheetRowStreamIteratorState{I,S}
    row_iter::I
    row_state::S
    rowcells::Dict{Int,Cell}
    local_formulas::Dict{SheetCellRef,AbstractFormula}
    rows_since_merge::Int
end

mutable struct WorksheetCacheIteratorState
    row_from_last_iteration::Int
end

mutable struct WorksheetCache{I<:SheetRowIterator} <: SheetRowIterator
    is_full::Bool # false until iterator runs to completion
    cells::CellCache # SheetRowNumber -> Dict{column_number, Cell}
    rows_in_cache::Vector{Int} # ordered vector with row numbers that are stored in cache
    row_ht::Dict{Int, Union{Float64, Nothing}} # Maps a row number to a row height
    row_index::Dict{Int, Int} # maps a row number to the index of the row number in rows_in_cache
    stream_iterator::I
    stream_state::Union{Nothing, SheetRowStreamIteratorState}
    dirty::Bool #indicate that data are not sorted, avoid sorting if we dont use the iterator
end

Base.isempty(wc::WorksheetCache) = isempty(wc.rows_in_cache)


# Excel Tables
struct TableStyleInfo
    name::Union{String,Nothing}
    show_first_column::Bool
    show_last_column::Bool
    show_row_stripes::Bool
    show_column_stripes::Bool
end

"""
    Table

A native Excel Table: a named, structured range within a worksheet, created in Excel
with *Insert → Table* (or `Ctrl+T`), or in XLSX.jl with [`XLSX.addtable!`](@ref).

Obtain one with [`XLSX.table`](@ref) or [`XLSX.tables`](@ref). A `Table` conforms to the
`Tables.jl` interface, so it can be passed directly to any compatible sink (e.g.
`DataFrame`), or read with [`XLSX.gettable`](@ref), [`XLSX.getdata`](@ref) or
[`XLSX.eachtablerow`](@ref). In every case only the Table's data rows are returned: the
header row and, if present, the totals row are excluded.

# Fields
- `id::Int`: the Table's id, unique within the workbook.
- `name::String`: the Table's name, unique within the workbook and sharing a namespace
  with defined names.
- `display_name::String`: the name Excel displays. Usually identical to `name`.
- `ref::CellRange`: the Table's full extent, including its header row and, if present,
  its totals row.
- `columns::Vector{String}`: the column names, in order, taken from the header row.
- `has_totals_row::Bool`: whether the last row of `ref` is a totals row.
- `style::Union{TableStyleInfo,Nothing}`: the Table's style, or `nothing` if it has none.
- `sheet`: the `Worksheet` the Table belongs to. Cell values are read from it on demand.

A `Table` is an immutable snapshot of the Table's structure at the time it was read.
Functions that modify a Table ([`XLSX.settotals!`](@ref), [`XLSX.appendtable!`](@ref))
return an updated `Table`; any earlier one goes stale and should be discarded.

See also [`XLSX.table`](@ref), [`XLSX.tables`](@ref), [`XLSX.addtable!`](@ref),
[`XLSX.deletetable!`](@ref), [`XLSX.settotals!`](@ref), [`XLSX.appendtable!`](@ref).
"""
struct Table
    id::Int
    name::String
    display_name::String
    ref::CellRange
    columns::Vector{String}
    has_totals_row::Bool
    style::Union{TableStyleInfo,Nothing}
    sheet # untyped to resolve circular dependency with Worksheet

end


"""
A `Worksheet` represents a reference to an Excel Worksheet.

From a `Worksheet` you can query for Cells, cell values and ranges.

# Example

```julia
xf = XLSX.readxlsx("myfile.xlsx")
sh = xf["mysheet"] # get a reference to a Worksheet
println( sh[2, 2] ) # access element "B2" (2nd row, 2nd column)
println( sh["B2"] ) # you can also use the cell name
println( sh["A2:B4"] ) # or a cell range
println( sh[:] ) # all data inside worksheet's dimension
```
"""
mutable struct Worksheet
    package::MSOfficePackage
    sheetId::Int
    relationship_id::String
    name::String
    dimension::Union{Nothing, CellRange}
    is_hidden::Bool
    cache::Union{WorksheetCache, Nothing}
    next_formula_id::Int
    unhandled_attributes::Union{Nothing,Dict{Int,Dict{String,String}}}
    sst_count::Int
    next_cf_priority::Union{Int, Nothing}
    tables_cache::Union{Vector{Table}, Nothing}   # nothing until first access to `tables(ws)`

    function Worksheet(package::MSOfficePackage, sheetId::Int, relationship_id::String, name::String, dimension::Union{Nothing, CellRange}, is_hidden::Bool)
        return new(package, sheetId, relationship_id, name, dimension, is_hidden, nothing, 0, nothing, 0, nothing, nothing)
    end
end

struct SheetRowStreamIterator <: SheetRowIterator
    sheet::Worksheet
end

#------------------------------------------------------------------------------ sharedStrings
mutable struct SharedStringTable
    shared_strings::Vector{String}
    unformatted::Vector{String}
    index::Dict{String, Int64} # for search optimisation. Tuple of indices to handle hash collisions.
    is_loaded::Bool
end

const ValidRichTextAttributes = [:bold, :italic, :under, :strike, :vertAlign, :color, :size, :name]

"""
    RichTextRun(text::String, pairs::Union{Nothing,Vector{Pair{Symbol,Any}}}=nothing)     -> RichTextRun
    RichTextRun(text::String)                                                             -> RichTextRun

Create an instance of a RichTextRun, representing a formatted substring element (run) to form part 
of a RichTextString. Each RichTextRun defines none, one or several font attributes to apply to its text.

- `text` specifies the text of the run's substring element.
- `pairs` is a vector of formatting attributes to apply to `text` (default = `nothing`).

Valid attributes that can be defined in `pairs` are:
- `:bold` - set `:bold => true` for this run to be emboldened. Omit otherwise.
- `:italic` - set `:italic => true` for this run to be italicised. Omit otherwise.
- `:under` - set `:under => true` to underline this run. Omit otherwise.
- `:strike` - set `:strike => true` to apply strikethrough to this run. Omit otherwise.
- `:vertAlign` - whether this run is `subscript` or `superscript` (eg `:vertAling => "superscript"`). Omit otherwise.
- `:color` - the color of this run (eg `:color => "red"`).
- `:size` - the size of the font to be used (eg `:size => 12`).
- `:name` - the name of the font to be used (e.g. `:name => "Arial"`).

Omit `pairs` to specify a run without formatting.

See also [`XLSX.RichTextString`](@ref).

# Examples
```julia
julia> rt1 = XLSX.RichTextRun("Water is H")
RichTextRun ("Water is H"  [ ])

julia> rt2 = XLSX.RichTextRun("2", [:vertAlign => "subscript"])
RichTextRun ("2"  [:vertAlign => "subscript"])

julia> rt3 = XLSX.RichTextRun("O!")
RichTextRun ("O!"  [ ])

julia> rt = XLSX.RichTextString(rt1, rt2, rt3)
RichTextString: "Water is H2O!" 
 containing 3 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Water is H"             [ ]
 "2"                      [:vertAlign => "subscript"]
 "O!"                     [ ]

julia> s["A1"] = rt
RichTextString: "Water is H2O!" 
 containing 3 runs:
 Run text                 Run attributes
 -------------------------------------------------------------------------------------------
 "Water is H"             [ ]
 "2"                      [:vertAlign => "subscript"]
 "O!"                     [ ]

```
![image|320x500](../images/H2O.png)
"""
struct RichTextRun
    text::String
    atts::Union{Nothing, Dict{Symbol,Any}}
    
    function RichTextRun(text::String, pairs::Union{Nothing,Vector{Pair{Symbol,Any}}}=nothing)
        isempty(text) && throw(XLSXError("Cannot create a RichTextRun with no text."))
        if isnothing(pairs)
            new(text, nothing)
        else
            atts=Dict(pairs)
            for x in keys(atts)
                in(x, ValidRichTextAttributes) || throw(XLSXError("Unknown Rich Text Attribute: ':$x'. Valid attributes are :bold, :italic, :under, :strike, :vertAlign, :color, :size, :name."))
            end
            new(text, atts)
        end
    end

end

"""
    RichTextString(runs::RichTextRun...)      -> RichTextString
    RichTextString(runs::Vector{RichTextRun}) -> RichTextString

Create an instance of a RichTextString from a set of RichTextRuns. A RichTextString supports rich text 
formatting within a single cell and is made up of multiple substrings (runs), each with different font 
attributes. The text in the cell is the simple concatenation of the text of each run but Excel will display 
each run with its own distinct font formatting within the cell. See also [`XLSX.RichTextRun`](@ref).

If a `RichTextString` containing only one run is assigned to a cell, the text will be assigned as plain 
text and the formatting attributes will be implemented on the whole cell using [`XLSX.setFont`](@ref).

See also [`XLSX.RichTextRun`](@ref).

# Examples
```julia
julia> rt = XLSX.RichTextString(rtf1, rtf2, rtf3, rtf4) # Create a RichTextString from four separate RichTextRuns.

julia> rt = XLSX.RichTextString([rtf1, rtf2, rtf3, rtf4]) # Create a RichTextString from a vector of four RichTextRuns.

```
"""
struct RichTextString <: AbstractString
    text::String
    runs::Vector{RichTextRun}

    function RichTextString(text::String, runs::Vector{RichTextRun })
        (isempty(text) || isempty(runs)) && throw(XLSXError("Cannot create an empty RichTextString"))
        new(text, runs)
    end
end

const DefinedNameValueTypes = Union{SheetCellRef, SheetCellRange, NonContiguousRange, CellConcreteType}#Int, Float64, String, Missing}
const DefinedNameRangeTypes = Union{SheetCellRef, SheetCellRange, NonContiguousRange}

struct DefinedNameValue
    value::DefinedNameValueTypes
    isabs::Union{Bool, Vector{Bool}}
    hidden::Bool
end

"""
    DefinedName

A defined name and its definition, as returned by [`getDefinedNames`](@ref) and
[`getAllDefinedNames`](@ref).

# Fields
- `name::String` — the defined name.
- `scope::Union{Nothing,String}` — `nothing` for a workbook-scoped name, or the
  display name of the worksheet it is scoped to.
- `value::DefinedNameValueTypes` — a `SheetCellRef`, `SheetCellRange` or
  `NonContiguousRange` for a name referring to a range, or the constant itself.
- `absolute::Union{Bool,Vector{Bool}}` — whether the reference is written as an
  absolute one (`\$A\$1` rather than `A1`); a vector, one entry per part, for a
  `NonContiguousRange`. Always `false` for a constant.
- `hidden::Bool` - whether the reference is hidden (system defined).

Together these are enough to recreate the name: addDefinedName(x, dn.name, dn.value; absolute=dn.absolute). 
This holds for the names getDefinedNames returns; names Excel generates for itself cannot be recreated, 
and are excluded from that result.

This is a snapshot of the definition at the time it was read, not a live handle:
it does not track later edits, and renaming a worksheet does not update the
`scope` or `value` of a `DefinedName` already in hand.
"""
struct DefinedName
    name::String
    scope::Union{Nothing,String}
    value::DefinedNameValueTypes
    absolute::Union{Bool,Vector{Bool}}
    hidden::Bool
end

# Workbook is the result of parsing file `xl/workbook.xml`.
# The `xl/workbook.xml` will need to be updated using the Workbook_names and 
# worksheet_names from here when a workbook is saved in case any new defined 
# names have been created.
mutable struct Workbook
    package::MSOfficePackage
    sheets::Vector{Worksheet}
    date1904::Bool
    relationships::Vector{Relationship}
    formulas::Dict{SheetCellRef, AbstractFormula}
    sst::SharedStringTable
    buffer_styles_is_float::Dict{Int, Bool}
    buffer_styles_is_datetime::Dict{Int, Bool}
    styles_lock::ReentrantLock
    formulas_lock::ReentrantLock
    sst_lock::ReentrantLock
    workbook_names::Dict{String, DefinedNameValue}
    # Keyed by (sheetId, name). NOT by localSheetId: the `localSheetId`
    # attribute of <definedName> is a zero-based index into <sheets>, so it
    # shifts whenever sheets are added, deleted or reordered. sheetId is stable
    # for the life of the sheet, so it is what we key on; the conversion to and
    # from localSheetId happens only at the read and write boundaries.
    worksheet_names::Dict{Tuple{Int, String}, DefinedNameValue}
    styles_xroot::Union{XML.Node, Nothing}
    num_style_index_cache::Dict{Int, CellDataFormat}
    theme_xroot::Union{XML.Node, Nothing}
    theme_colors::Union{Vector{String}, Nothing}
    theme_color_map::Union{Nothing,Dict{String,String}}
    theme_font_map::Union{Nothing,Dict{String,String}}
    cellXfs_cache::Union{Vector{XML.Node}, Nothing}   # cache for get_cellXfs_nodes
    numFmt_cache::Union{Dict{Int, String}, Nothing}   # cache for get_numFmt_cache
    style_table_cache::Dict{String, Vector{XML.Node}} # cache for fonts/borders/fills, keyed by tag ("fonts","borders","fills")
    next_table_id::Union{Int,Nothing}   # nothing until first computed
end

@enum TemplateType begin
    NotATemplate
    XLTXTemplate   # .xltx → save as .xlsx
    XLTMTemplate   # .xltm → save as .xlsm
end

"""
`XLSXFile` represents a reference to an Excel file.

It is created by using [`XLSX.readxlsx`](@ref), [`XLSX.openxlsx`](@ref), 
[`XLSX.opentemplate`](@ref) or [`XLSX.newxlsx`](@ref).

From an `XLSXFile` you can navigate to an `XLSX.Worksheet` reference
as shown in the example below.

# Example

```julia
xf = XLSX.readxlsx("myfile.xlsx")
sh = xf["mysheet"] # get a reference to a Worksheet
```
"""
mutable struct XLSXFile <: MSOfficePackage
    source::Union{AbstractString, IO}
    use_cache_for_sheet_data::Bool
    load_formulas::Bool                          # ← new
    files::Dict{String, Bool}
    data::Dict{String, Union{XML.Node, String}}
    namespace::Dict{String, Union{String, Nothing}}
    binary_data::Dict{String, Vector{UInt8}}
    workbook::Workbook
    relationships::Vector{Relationship}
    is_writable::Bool
    template_type::TemplateType
    uuid_rng::Random.Xoshiro

    function XLSXFile(source::Union{AbstractString, IO}, use_cache::Bool, is_writable::Bool, load_formulas::Bool=true)
        check_for_xlsx_file_format(source)
        xl = new(
            source,
            use_cache,
            load_formulas,
            Dict{String, Bool}(),
            Dict{String, Union{XML.Node, String}}(),
            Dict{String, Union{String, Nothing}}(),
            Dict{String, Vector{UInt8}}(),
            EmptyWorkbook(),
            Vector{Relationship}(),
            is_writable,
            NotATemplate,
            Random.Xoshiro(2468)
        )
        xl.workbook.package = xl
        return xl
    end
end

struct ReadFile
    node::Union{Nothing,XML.Node,String}
    raw::Union{Nothing,String}
    bin::Union{Nothing,Vector{UInt8}}
    name::String
end

#
# Iterators
#

struct SheetRow
    sheet::Worksheet
    row::Int                     # index of the row in the worksheet
    ht::Union{Float64, Nothing}  # row height
    rowcells::Dict{Int, Cell}    # column -> value
end

struct Index # for TableRowIterator - based on DataFrames.jl
    lookup::Dict{Symbol, Int} # column label -> table column index
    column_labels::Vector{Symbol}
    column_map::Dict{Int, Int} # table column index (1-based) -> sheet column index (cellref based)

    function Index(column_range::Union{ColumnRange, AbstractString}, column_labels)
        column_labels_as_syms = [ Symbol(i) for i in column_labels ]
        column_range = convert(ColumnRange, column_range)
        if length(unique(column_labels_as_syms)) != length(column_labels_as_syms)
            throw(XLSXError("Column labels must be unique."))
        end

        lookup = Dict{Symbol, Int}()
        for (i, n) in enumerate(column_labels_as_syms)
            lookup[n] = i
        end

        column_map = Dict{Int, Int}()
        for (i, n) in enumerate(column_range)
            column_map[i] = decode_column_number(n)
        end
        return new(lookup, column_labels_as_syms, column_map)
    end
end

struct TableRowIterator{I<:SheetRowIterator}
    itr::I
    index::Index
    first_data_row::Int
    stop_in_empty_row::Bool
    stop_in_row_function::Union{Nothing, Function}
    keep_empty_rows::Bool
    missing_strings::Set{String}
    resume::Union{Nothing, Tuple{SheetRow, Any}}  # pre-fetched (row, state) to start from, or nothing
end

struct TableRow
    row::Int # Index of the row in the table. This is not relative to the worksheet cell row.
    index::Index
    cell_values::Vector{CellConcreteType}
end

struct TableRowIteratorState{S}
    table_row_index::Int
    sheet_row_index::Int
    sheet_row_iterator_state::S
    missing_rows::Int # number of completely empty rows between the last row and the current row
    row_pending::Union{Nothing, SheetRow} # if the last row was empty, this is the row that was pending to be returned
end

struct XLSXTableRow
    table::Table
    row_number::Int
end

struct XLSXTableRowIterator
    table::Table
end

"""
`XLSX.DataTable` is a simple `Tables.jl` compatibledata structure to hold tabular 
data extracted from an Excel worksheet.

It is created by [`XLSX.gettable`](@ref) or, direct from a file, with [`XLSX.readtable`](@ref).

Pass a `DataTable` to any `Tables.jl` compatible sink, e.g. `DataFrame(dt)`.

"""
struct DataTable
    data::Vector{Any} # columns
    column_labels::Vector{Symbol}
    column_label_index::Dict{Symbol, Int} # column_label -> column_index

    function DataTable(
            data::Vector{Any}, # columns
            column_labels::Vector{Symbol},
        )

        if length(data) != length(column_labels)
            throw(XLSXError("Data has $(length(data)) columns but $(length(column_labels)) column labels."))
        end

        column_label_index = Dict{Symbol, Int}()
        for (i, sym) in enumerate(column_labels)
            if haskey(column_label_index, sym)
                throw(XLSXError("DataTable has repeated label for column `$sym`"))
            end
            column_label_index[sym] = i
        end

        return new(data, column_labels, column_label_index)
    end
end

struct XPathInfo
    node::XML.Node
    path::String

    function XPathInfo(node::XML.Node, path::String)
        new(node, path)
    end
end

struct XLSXError <: Exception
    msg::String
end
Base.showerror(io::IO, e::XLSXError) = print(io, "XLSXError: ",e.msg)

struct FileArray <: AbstractVector{UInt8}
    filename::String
    offset::Int64
    len::Int64
end

mutable struct Locked{T}
    value::T
    lock::ReentrantLock
    Locked(x::T) where {T} = new{T}(x, ReentrantLock())
end

#=
function withlock(f, obj::Locked)
    lock(obj.lock) do
        f(obj.value)
    end
end
=#

# ===========================================================================
# Charts
# ===========================================================================

"""
`ChartRef`

One cached reference from a chart series: the formula it came from, the number
format Excel recorded for it, and the cached values themselves.

# Fields
- `kind::Symbol` - one of `:num`, `:str`, `:multiLvlStr`, `:numLit`, `:strLit`.
- `ref::Union{Nothing,String}` - the `c:f` formula (`Sheet1!\$B\$2:\$B\$9`).
  `nothing` for literal (`c:numLit` / `c:strLit`) series, which have no source range.
- `format_code::Union{Nothing,String}` - number format recorded in the cache.
- `ptCount::Int` - number of points Excel declared, whether or not the cache was read.
- `data::Vector` - cached values, length `ptCount`, gaps as `missing`. Empty when
  the chart was read with `cache=false`.
- `errors::Dict{Int,UInt64}` - index => error code for cached error values.

Excel only caches the error values `#N/A` in the chart data cache. Others are written 
a 0 and become indistinguishable from real zero in the chart cache.

`ChartRef` only applies to `c:` charts and not `cx:` charts.


!!! note
    For `kind == :multiLvlStr` each element of `data` is itself a level vector,
    in document order (Excel writes the innermost/leaf level first).
"""
struct ChartRef
    kind::Symbol
    ref::Union{Nothing,String}
    format_code::Union{Nothing,String}
    ptCount::Int
    data::Vector
    errors::Dict{Int,UInt64}
end

# chart part path => (sheet name, from, to, rId)
const ChartAnchor = NamedTuple{
    (:sheet, :from, :to, :rId),
    Tuple{String,Union{Nothing,String},Union{Nothing,String},String},
}

# Internal: one discovered chart part, before parsing.
struct ChartLocation
    path::String
    anchor::Union{Nothing,ChartAnchor}
    schema::Symbol   # :c or :cx
end

"""
    ChartSeries

One series (`c:ser`), as read.

`categories` holds `c:cat` for category charts and `c:xVal` for scatter and bubble
charts; `values` holds `c:val` or `c:yVal` correspondingly, so the two fields mean
the same thing whatever the chart type.

Identified by `idx`, the series' `c:idx`, which is unique within the chart and
unaffected by adding series. `order` is the plotting order (`c:order`) and is not
a key. Functions that take a series position `i` resolve it against the current
part on each call. The other fields describe the series as it was when read.
`raw` is the `c:ser` element at that time, or `nothing` for a series built in
code; it is never used to address the part.
"""
struct ChartSeries
    idx::Int
    order::Int
    charttype::Symbol
    name::Union{Nothing,String}
    name_ref::Union{Nothing,ChartRef}
    categories::Union{Nothing,ChartRef}
    values::Union{Nothing,ChartRef}
    bubble_sizes::Union{Nothing,ChartRef}
    raw::Union{Nothing,XML.Node}          # the c:ser element
end

const ChartRange = Union{Nothing,SheetCellRef,SheetCellRange,SheetRowRange,SheetColumnRange,NonContiguousRange}

const ChartRanges = @NamedTuple{
    idx::Int,
    name::Union{Nothing,String},
    categories::ChartRange,
    values::ChartRange,
    bubble_sizes::ChartRange,
}


"""
    AbstractChart

Supertype for charts read from a workbook. Two concrete subtypes exist:
[`Chart`](@ref) for the standard `c:` schema, and [`ChartEx`](@ref) for the
newer `cx:` schema used by waterfall, funnel, treemap, sunburst, histogram,
Pareto, box & whisker and region map charts.
"""
abstract type AbstractChart end

"""
    Chart

A handle to one `c:` chart part. It holds the part's identity and anchor only;
everything else is read from the current part on each call, so a `Chart` never
goes stale and setters return it unchanged.

# Fields
- `package` — the `XLSXFile` the part belongs to.
- `path` — package path, e.g. `"xl/charts/chart1.xml"`.
- `name` — part name without extension, e.g. `"chart1"`.
- `rId` — relationship id of the chart within its drawing part, if resolved.
- `sheet` — name of the sheet the chart is anchored to, if resolved.
- `from`, `to` — anchor cell references as strings, following `getImages`.

Content is reached through accessors: [`getChartTitle`](@ref),
[`getChartTypes`](@ref), [`getChartSeries`](@ref), [`getChartData`](@ref) and the
rest of the chart API. The values they return carry keys, so they remain valid
arguments after later writes; see [`ChartSeries`](@ref).
"""
struct Chart <: AbstractChart
    package::XLSXFile
    path::String
    name::String
    rId::Union{Nothing,String}
    sheet::Union{Nothing,String}
    from::Union{Nothing,String}
    to::Union{Nothing,String}
end

"""
    ChartEx

A handle to one chartEx part: a chart in the Microsoft `cx:` namespace
(`http://schemas.microsoft.com/office/drawing/2014/chartex`), used by Excel for
waterfall, funnel, treemap, sunburst, histogram, Pareto, box & whisker and
region map charts. Charts in the ECMA-376 `c:` namespace are [`Chart`](@ref).

A `ChartEx` holds only the part's identity and its anchor. Its content (layout,
data references, title) is read from the part on each call, so a `ChartEx`
remains valid after the part is written to.

# Fields
- `package` - the `XLSXFile` containing the chart.
- `path` - package path, e.g. `"xl/charts/chartEx1.xml"`.
- `name` - part name without extension, e.g. `"chartEx1"`.
- `rId` - relationship id of the chart within its drawing part, if resolved.
- `sheet` - name of the sheet the chart is anchored to, if resolved.
- `from`, `to` - anchor cell references as strings, following `getImages`.

# Reading content
- [`chartType`](@ref) - e.g. `:waterfall`, `:histogram`.
- [`getChartTitle`](@ref) - title text, whether typed or bound to a cell;
  `nothing` if the title has no text of its own.
- [`getChartRanges`](@ref) - the source range of each data dimension, resolved
  through the workbook's hidden `_xlchart.*` defined names.

See also [`getCharts`](@ref), [`chartSchema`](@ref).
"""
struct ChartEx <: AbstractChart
    package::XLSXFile
    path::String
    name::String
    rId::Union{Nothing,String}
    sheet::Union{Nothing,String}
    from::Union{Nothing,String}
    to::Union{Nothing,String}
end

"""
    ChartExDimension

One data dimension of a chartEx data block: a `cx:numDim` or `cx:strDim`.

# Fields
- `kind` - `:num` or `:str`.
- `type` - the dimension's role as written, e.g. `:val`, `:cat`, `:size`,
  `:x`, `:y`, `:colorVal`, `:colorStr`, `:entityId`.
- `formula` - the text of `cx:f`, usually a hidden `_xlchart.*` defined name;
  `nothing` if the dimension has no formula.
- `range` - the resolved source range; `nothing` if unresolved.
"""
struct ChartExDimension
    kind::Symbol
    type::Symbol
    formula::Union{Nothing,String}
    range::ChartRange
end

"""
    ChartExData

One `cx:data` block. Series refer to a block by `id` through `cx:dataId`.
"""
struct ChartExData
    id::Int
    dimensions::Vector{ChartExDimension}
end

"""
    ChartExBinning

Histogram binning from `cx:layoutPr/cx:binning`. Field names follow the XML.
A numeric field may also be `:auto`; `nothing` means not written.

# Fields
- `intervalClosed` - `:r` or `:l`, the closed side of each bin interval.
- `underflow`, `overflow` - the underflow and overflow bin cut-offs.
- `binSize` - bin width.
- `binCount` - number of bins.
"""
struct ChartExBinning
    intervalClosed::Union{Nothing,Symbol}
    underflow::Union{Nothing,Float64,Symbol}
    overflow::Union{Nothing,Float64,Symbol}
    binSize::Union{Nothing,Float64,Symbol}
    binCount::Union{Nothing,Int,Symbol}
end

Base.:(==)(a::ChartExDimension, b::ChartExDimension) =
    a.kind == b.kind && a.type == b.type && a.formula == b.formula && a.range == b.range
Base.hash(d::ChartExDimension, h::UInt) =
    hash((d.kind, d.type, d.formula, d.range), hash(:ChartExDimension, h))

Base.:(==)(a::ChartExData, b::ChartExData) = a.id == b.id && a.dimensions == b.dimensions
Base.hash(d::ChartExData, h::UInt) = hash((d.id, d.dimensions), hash(:ChartExData, h))

Base.:(==)(a::ChartExBinning, b::ChartExBinning) =
    all(getfield(a, f) == getfield(b, f) for f in fieldnames(ChartExBinning))
Base.hash(b::ChartExBinning, h::UInt) =
    hash(Tuple(getfield(b, f) for f in fieldnames(ChartExBinning)), hash(:ChartExBinning, h))

const SCHEME_TOKENS = (:bg1, :tx1, :bg2, :tx2, :accent1, :accent2, :accent3,
                       :accent4, :accent5, :accent6, :hlink, :folHlink,
                       :lt1, :dk1, :lt2, :dk2)

const COLOR_TRANSFORMS = (:lumMod, :lumOff, :shade, :tint, :satMod, :alpha)

# bg1/lt1, tx1/dk1, bg2/lt2 and tx2/dk2 name the same slot. The field keeps the
# token as written so a round trip preserves it; equality and hashing normalize
# so two spellings of one color compare equal.
const SCHEME_ALIASES = Dict(:lt1 => :bg1, :dk1 => :tx1, :lt2 => :bg2, :dk2 => :tx2)

"""
    SchemeColor(token; lumMod = nothing, lumOff = nothing, ...)
    SchemeColor(token, transforms)

A DrawingML theme color: a token naming a slot in the workbook's color scheme,
plus any transforms applied to it. Distinct from the spreadsheet
`<color theme="N"/>` mechanism, which is index-ordered and uses a different
tint algorithm — see `get_theme_colors` for that.

`token` is one of `:bg1`, `:tx1`, `:bg2`, `:tx2`, `:accent1` … `:accent6`,
`:hlink`, `:folHlink`, or the aliases `:lt1`, `:dk1`, `:lt2`, `:dk2`.

The `:lt1`, `:dk1`, `:lt2` and `:dk2` aliases are preserved as written, so a
round trip keeps the spelling the file used, but two spellings of one slot
compare and hash equal.

Transforms are held in document order, because DrawingML applies them in
sequence and `lumMod` before `lumOff` is not the same as the reverse. Equality
is positional for the same reason: two `SchemeColor`s with the same transforms
in different orders are not equal, because they do not render the same. Values
are percentages, not the hundredths DrawingML writes: `lumMod = 75` is 75%.

The keyword constructor emits transforms in the conventional order — `lumMod`
before `lumOff`, shade or tint last. Any other order needs the vector form.

# Examples

    SchemeColor(:accent1)
    SchemeColor(:accent1; lumMod = 75)                    # 104862 against the Office theme
    SchemeColor(:tx1; lumMod = 65, lumOff = 35)           # 595959
    SchemeColor(:accent2, [:shade => 50.0, :alpha => 80.0])
"""
struct SchemeColor
    token::Symbol
    transforms::Vector{Pair{Symbol,Float64}}

    function SchemeColor(token::Symbol, transforms::Vector{Pair{Symbol,Float64}})
        token in SCHEME_TOKENS || throw(XLSXError(
            "`$token` is not a theme color token. Valid tokens: " *
            join(SCHEME_TOKENS, ", ") * "."))
        for (name, _) in transforms
            name in COLOR_TRANSFORMS || throw(XLSXError(
                "`$name` is not a color transform. Valid transforms: " *
                join(COLOR_TRANSFORMS, ", ") * "."))
        end
        return new(token, transforms)
    end
end

function SchemeColor(token::Symbol; lumMod = nothing, lumOff = nothing,
                     satMod = nothing, shade = nothing, tint = nothing,
                     alpha = nothing)
    t = Pair{Symbol,Float64}[]
    isnothing(lumMod) || push!(t, :lumMod => Float64(lumMod))
    isnothing(lumOff) || push!(t, :lumOff => Float64(lumOff))
    isnothing(satMod) || push!(t, :satMod => Float64(satMod))
    isnothing(shade)  || push!(t, :shade  => Float64(shade))
    isnothing(tint)   || push!(t, :tint   => Float64(tint))
    isnothing(alpha)  || push!(t, :alpha  => Float64(alpha))
    return SchemeColor(token, t)
end

canonical_token(t::Symbol) = get(SCHEME_ALIASES, t, t)

Base.:(==)(a::SchemeColor, b::SchemeColor) =
    canonical_token(a.token) == canonical_token(b.token) && a.transforms == b.transforms

Base.hash(c::SchemeColor, h::UInt) =
    hash(c.transforms, hash(canonical_token(c.token), hash(:SchemeColor, h)))

Base.isequal(a::SchemeColor, b::SchemeColor) = a == b

"""
    DrawingColor

A DrawingML colour: the element as written, plus the RGB it resolves to.

`kind` and `val` record the reference as authored - a theme colour stays a
theme colour - and `transforms` the modifications applied to it, in document
order. `rgb` and `alpha` give the resolved result for anyone who just wants to
know what it looks like.

# Fields
- `kind::Symbol` - `:srgb`, `:scheme`, `:sys`, `:prst`, `:hsl` or `:scrgb`.
- `val::String` - the `val` attribute: `"FF0000"`, `"accent1"`, `"windowText"`.
- `transforms::Vector{Pair{Symbol,Int}}` - e.g. `[:lumMod => 60000, :lumOff => 40000]`,
  in thousandths of a percent, in the order DrawingML applies them.
- `rgb::String` - the resolved colour as `"RRGGBB"`.
- `alpha::Float64` - `1.0` unless an `alpha` transform applies.

`rgb` and `alpha` are resolved values, not source, and are not written — the
file keeps the scheme reference and its transforms so the theme still applies.
"""
struct DrawingColor
    kind::Symbol
    val::String
    transforms::Vector{Pair{Symbol,Int}}
    rgb::String
    alpha::Float64
end

"""
    ColorSpec

Anything that can specify a color where one is written: a parsed
[`DrawingColor`](@ref), a [`SchemeColor`](@ref), or a string naming an
`AARRGGBB` value or a Colors.jl color.

The parser only ever produces `DrawingColor`, so a value read from a file is
always that. The other two exist for construction.
"""
const ColorSpec = Union{DrawingColor, SchemeColor, String}

"""
    DrawingFill

A DrawingML fill: `<a:solidFill>`, `<a:noFill>`, `<a:gradFill>`, `<a:pattFill>`
or `<a:blipFill>`.

A solid fill resolves to one colour in `fgcolor`. A pattern fill resolves to
two, `fgcolor` and `bgcolor`, with the pattern itself in `preset` - matching
how cell fills are exposed. Gradient and picture fills are identified by `kind`
but not modelled further: `raw` holds the element as read, so nothing is lost
on write.

# Fields
- `kind::Symbol` - `:none`, `:solid`, `:gradient`, `:pattern`, `:blip` or `:group`.
- `fgcolor` - the colour of a solid fill, or a pattern's foreground.
- `bgcolor` - a pattern's background; `nothing` otherwise.
- `preset::Union{Nothing,String}` - a pattern's `prst` attribute, e.g. `"pct25"`,
  `"ltUpDiag"`.
- `raw::Union{Nothing,XML.Node}` - the element as read, or `nothing` for a fill
  built in code. Gradient, pattern, picture and group fills are written back from
  it.
"""
struct DrawingFill
    kind::Symbol
    fgcolor::Union{Nothing,ColorSpec}
    bgcolor::Union{Nothing,ColorSpec}
    preset::Union{Nothing,String}
    raw::Union{Nothing,XML.Node}
end

"""
    DrawingFill(kind; fgcolor = nothing, bgcolor = nothing, preset = nothing)

A fill built rather than parsed. `kind` is `:none`, `:solid`, `:gradient`,
`:pattern`, `:blip` or `:group`; only `:none` and `:solid` can be serialized
from a constructed value, since the others are modelled partially and written
back from `raw`.
A fill read from a file carries a DrawingColor with `rgb` resolved; one built
by hand may carry a SchemeColor or a string, which has no resolved value until
it is written and read back.
"""
DrawingFill(kind::Symbol; fgcolor = nothing, bgcolor = nothing, preset = nothing) =
    DrawingFill(kind, fgcolor, bgcolor, preset, nothing)

"""
    DrawingLine

A DrawingML outline: `<a:ln>`.

The stroke colour lives in `fill`, since a line is filled the same way a shape
is - solid, gradient, pattern or none.

# Fields
- `fill::Union{Nothing,DrawingFill}` - the stroke; `nothing` where the element
  says nothing about it, `kind === :none` where it explicitly has no outline.
- `width::Union{Nothing,Float64}` - the `w` attribute, in points (the file stores
  EMU, 12700 to the point).
- `dash::Union{Nothing,String}` - `"solid"`, `"dash"`, `"sysDot"`, and so on.
- `cap::Union{Nothing,String}` - `"rnd"`, `"sq"`, `"flat"`.
- `compound::Union{Nothing,String}` - the `cmpd` attribute: `"sng"`, `"dbl"`, …
- `join::Union{Nothing,String}` - `"round"`, `"bevel"` or `"miter"`, from the
  `a:round`/`a:bevel`/`a:miter` child.
- `miter_limit::Union{Nothing,Float64}` - the `a:miter` `lim`, as a fraction of
  the line width; only with a miter join.
- `raw::Union{Nothing,XML.Node}` - the element as read, or `nothing` for a line
  built in code.
"""
struct DrawingLine
    fill::Union{Nothing,DrawingFill}
    width::Union{Nothing,Float64}         # w, points (file stores EMU)
    dash::Union{Nothing,String}
    cap::Union{Nothing,String}
    compound::Union{Nothing,String}
    join::Union{Nothing,String}           # :round, :bevel, :miter
    miter_limit::Union{Nothing,Float64}   # fraction of line width; only with miter
    raw::Union{Nothing,XML.Node}
end

DrawingLine(; fill = nothing, width = nothing, dash = nothing, cap = nothing,
              compound = nothing, join = nothing, miter_limit = nothing) =
    DrawingLine(fill, width, dash, cap, compound, join, miter_limit, nothing)

"""
    DrawingShapeProps

Shape properties (`a:spPr`) — the fill and outline of anything drawn in a chart:
series, data points, the plot area, the chart area, axis lines, legend, gridlines.

`fill` and `line` distinguish three states, and the difference matters:

| file                            | reads as               | means                  |
|---------------------------------|------------------------|------------------------|
| no fill element                 | `nothing`              | inherit from the style |
| `<a:noFill/>`                   | `kind == :none`        | deliberately invisible |
| `<a:solidFill>…`                | `kind == :solid`       | this colour            |

The same applies to `line.fill`: `<a:ln><a:noFill/></a:ln>` is how Excel writes
"no border", which is not the same as omitting `a:ln` entirely.

`effects` holds `a:effectLst` or `a:effectDag` as an unparsed node — enough to
report that a shape has effects without modelling shadows and glows. Geometry
(`a:xfrm`, `a:prstGeom`, `a:custGeom`) and 3-D (`a:scene3d`, `a:sp3d`) are not
modelled at all; they stay in `raw`, which is where chart creation (stage 6)
will find them. Chart parts rarely carry geometry — it belongs to the drawing
shapes that host the chart, not the chart itself.
"""
struct DrawingShapeProps
    fill::Union{Nothing,DrawingFill}
    line::Union{Nothing,DrawingLine}
    effects::Union{Nothing,XML.Node}    # a:effectLst or a:effectDag
    bwmode::Union{Nothing,String}       # bwMode: clr | auto | gray | ltGray | invGray | ...
    raw::Union{Nothing,XML.Node}
end

# =============================================================================
# Every optional field is Union{Nothing,T}: absent means "inherit from the
# theme or the parent list style", which is not the same as an explicit value,
# and the difference has to survive a round trip.
# =============================================================================

"""
    DrawingRunProps

Character-level properties: `a:rPr`, `a:defRPr` or `a:endParaRPr`.

Sizes are points (the file stores 1/100 pt), `baseline` is a percentage, and
`under` / `strike` / `caps` keep the DrawingML vocabulary as written
(`"sng"`, `"noStrike"`, `"small"`). Typefaces may be theme references —
`"+mn-lt"` for the minor latin font, `"+mj-lt"` for major.
"""
struct DrawingRunProps
    lang::Union{Nothing,String}
    size::Union{Nothing,Float64}        # sz, points
    bold::Union{Nothing,Bool}           # b
    italic::Union{Nothing,Bool}         # i
    under::Union{Nothing,String}        # u
    strike::Union{Nothing,String}
    caps::Union{Nothing,String}         # cap
    baseline::Union{Nothing,Float64}    # fraction of the font size
    kern::Union{Nothing,Float64}        # points
    spacing::Union{Nothing,Float64}     # spc, points
    fill::Union{Nothing,DrawingFill}
    line::Union{Nothing,DrawingLine}    # a:ln — text outline
    latin::Union{Nothing,String}
    ea::Union{Nothing,String}
    cs::Union{Nothing,String}
    raw::Union{Nothing,XML.Node}
end

DrawingRunProps(; lang = nothing, size = nothing, bold = nothing, italic = nothing,
                  under = nothing, strike = nothing, caps = nothing,
                  baseline = nothing, kern = nothing, spacing = nothing,
                  fill = nothing, line = nothing,
                  latin = nothing, ea = nothing, cs = nothing) =
    DrawingRunProps(lang, size, bold, italic, under, strike, caps, baseline,
                    kern, spacing, fill, line, latin, ea, cs, nothing)

"""
    DrawingParaProps

Paragraph properties (`a:pPr`). Margins and indent are points; spacing fields
are `(:pct, percent)` or `(:pts, points)`. `defprops` is the nested `a:defRPr`,
which in a chart `txPr` is usually the only place the font is specified.
"""
struct DrawingParaProps
    align::Union{Nothing,String}        # algn
    level::Union{Nothing,Int}           # lvl
    marginleft::Union{Nothing,Float64}  # marL, points
    marginright::Union{Nothing,Float64} # marR, points
    indent::Union{Nothing,Float64}      # points
    rtl::Union{Nothing,Bool}
    linespacing::Union{Nothing,Tuple{Symbol,Float64}}   # lnSpc
    spacebefore::Union{Nothing,Tuple{Symbol,Float64}}   # spcBef
    spaceafter::Union{Nothing,Tuple{Symbol,Float64}}    # spcAft
    defprops::Union{Nothing,DrawingRunProps}
    raw::Union{Nothing,XML.Node}
end

DrawingParaProps(; align = nothing, level = nothing,
                   marginleft = nothing, marginright = nothing, indent = nothing,
                   rtl = nothing, linespacing = nothing,
                   spacebefore = nothing, spaceafter = nothing,
                   defprops = nothing) =
    DrawingParaProps(align, level, marginleft, marginright, indent, rtl,
                     linespacing, spacebefore, spaceafter, defprops, nothing)


"""
    DrawingRun

One `a:r`, `a:br` or `a:fld`, distinguished by `kind` (`:run`, `:br`, `:fld`).
A break carries `"\\n"` as its text so `text_content` needs no special case.
"""
struct DrawingRun
    kind::Symbol
    text::Union{Nothing,String}
    props::Union{Nothing,DrawingRunProps}
    raw::Union{Nothing,XML.Node}
end

"""
    DrawingRun(text; props = nothing, kind = :run)

One run of text. `kind` is `:run`, `:br` or `:fld`; a break carries `"\\n"` as
its text.
"""
DrawingRun(text::AbstractString; props = nothing, kind::Symbol = :run) =
    DrawingRun(kind, String(text), props, nothing)


"""
    DrawingParagraph

One `a:p`: properties, runs in document order, and the trailing
`a:endParaRPr`, which is what Excel writes when a paragraph has no runs.
"""
struct DrawingParagraph
    props::Union{Nothing,DrawingParaProps}
    runs::Vector{DrawingRun}
    endprops::Union{Nothing,DrawingRunProps}
    raw::Union{Nothing,XML.Node}
end

"""
    DrawingParagraph(runs...; props = nothing, endprops = nothing)

One paragraph. Runs may be `DrawingRun`s or plain strings, which become runs
with no properties of their own — they inherit the paragraph default.
"""
DrawingParagraph(runs::Union{DrawingRun,AbstractString}...;
                 props = nothing, endprops = nothing) =
    DrawingParagraph(props,
                     DrawingRun[r isa DrawingRun ? r : DrawingRun(r) for r in runs],
                     endprops, nothing)

"""
    DrawingBodyProps

Text-body properties (`a:bodyPr`). `rotation` is degrees (stored as 1/60000),
insets are points, and `autofit` is `:none`, `:normal` or `:shape` — parsed
from a child element, with `fontscale` and `linespacereduction` populated only
for `:normal`.
"""
struct DrawingBodyProps
    rotation::Union{Nothing,Float64}           # rot, degrees
    vertical::Union{Nothing,String}            # vert
    wrap::Union{Nothing,String}
    anchor::Union{Nothing,String}
    anchorctr::Union{Nothing,Bool}
    upright::Union{Nothing,Bool}
    spcfirstlastpara::Union{Nothing,Bool}
    vertoverflow::Union{Nothing,String}
    horzoverflow::Union{Nothing,String}
    insetleft::Union{Nothing,Float64}
    insettop::Union{Nothing,Float64}
    insetright::Union{Nothing,Float64}
    insetbottom::Union{Nothing,Float64}
    autofit::Union{Nothing,Symbol}
    fontscale::Union{Nothing,Float64}          # fraction of the font size
    linespacereduction::Union{Nothing,Float64} # fraction of the font size
    raw::Union{Nothing,XML.Node}
end

DrawingBodyProps(; rotation = nothing, vertical = nothing, wrap = nothing,
                   anchor = nothing, anchorctr = nothing, upright = nothing,
                   spcfirstlastpara = nothing, vertoverflow = nothing,
                   horzoverflow = nothing,
                   insetleft = nothing, insettop = nothing,
                   insetright = nothing, insetbottom = nothing,
                   autofit = nothing, fontscale = nothing,
                   linespacereduction = nothing) =
    DrawingBodyProps(rotation, vertical, wrap, anchor, anchorctr, upright,
                     spcfirstlastpara, vertoverflow, horzoverflow,
                     insetleft, insettop, insetright, insetbottom,
                     autofit, fontscale, linespacereduction, nothing)

"""
    DrawingText

A DrawingML text body (`CT_TextBody`): `c:txPr`, `c:rich`, or the cx: equivalent.

`liststyle` is kept as an unparsed node — it is empty in most chart parts and
carries list-level defaults we don't model. `raw` is the text body element
itself, for surgical write-back.
"""
struct DrawingText
    body::Union{Nothing,DrawingBodyProps}
    liststyle::Union{Nothing,XML.Node}
    paragraphs::Vector{DrawingParagraph}
    raw::Union{Nothing,XML.Node}
end

"""
    DrawingText(paragraphs...; body = nothing, liststyle = nothing)

A text body. Paragraphs may be `DrawingParagraph`s or plain strings, each
becoming a one-run paragraph.

    DrawingText("Revenue by Region")
    DrawingText(DrawingParagraph("Revenue", DrawingRun(" 2026",
                    props = DrawingRunProps(bold = true))))
"""
DrawingText(paras::Union{DrawingParagraph,AbstractString}...;
            body = nothing, liststyle = nothing) =
    DrawingText(body, liststyle,
                DrawingParagraph[p isa DrawingParagraph ? p : DrawingParagraph(p)
                                 for p in paras],
                nothing)

"""
    ChartMarker

Marker properties (`c:marker` under a `c:ser` or `c:dPt`), as read.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `point_idx` — `c:idx` of the owning data point, or `nothing` for the series' own
  marker.
- `symbol` — `:circle`, `:square`, `:none`, …; `nothing` means absent.
- `size` — points, 2 to 72; `nothing` means absent.
- `shape` — the marker's `c:spPr`, if written.
- `raw` — the element as read, or `nothing` for a marker built in code.

Identified by `series_idx` and `point_idx`. `raw` is never used to address the part.
"""
struct ChartMarker
    series_idx::Int
    point_idx::Union{Nothing,Int}        # nothing for the series' own marker
    symbol::Union{Nothing,Symbol}
    size::Union{Nothing,Int}
    shape::Union{Nothing,DrawingShapeProps}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartAxis

One axis (`c:catAx`, `c:valAx`, `c:dateAx` or `c:serAx`), as read.

# Fields
- `kind` — the element tag as a Symbol, e.g. `:valAx`.
- `axid` — `c:axId`, the key.
- `pos` — `c:axPos`: `:b`, `:t`, `:l` or `:r`.
- `crossax` — `c:crossAx`, the `axid` of the axis this one crosses.
- `deleted` — `c:delete`; a deleted axis is not drawn but remains formattable.
- `raw` — the element as read, or `nothing` for an axis built in code.

Identified by `axid`. Functions taking `(c, ax)` find the axis by it in the
current part, so `ax` stays usable across writes. The other fields describe the
axis as it was when read; getters that depend on the axis kind check the current
element, since Excel keeps the `axId` when a category axis is switched to a date
axis. `raw` is never used to address the part.
"""
struct ChartAxis
    kind::Symbol
    axid::Union{Nothing,Int}
    pos::Union{Nothing,Symbol}
    crossax::Union{Nothing,Int}
    deleted::Bool
    raw::Union{Nothing,XML.Node}
end

"""
    ChartGroup

One chart-type group in `c:plotArea` (`c:barChart`, `c:lineChart`, …), as read.

# Fields
- `kind` — the element tag as a Symbol, e.g. `:barChart`.
- `axids` — the `c:axId` values the group plots against, in order; empty for pie
  and doughnut groups.
- `raw` — the element as read, or `nothing` for a group built in code.

Identified by `(kind, axids)`: a combo chart's groups plot against different axis
pairs, so the pair is unique in practice. Functions taking `(c, g)` find the
group by it in the current part and throw if no group, or more than one, matches.
Equality and hashing use the key alone. `raw` is never used to address the part.
"""
struct ChartGroup
    kind::Symbol
    axids::Vector{Int}
    raw::Union{Nothing,XML.Node}
end

# A group's identity is its key, so equality ignores `raw`.
Base.:(==)(a::ChartGroup, b::ChartGroup) = a.kind == b.kind && a.axids == b.axids
Base.hash(g::ChartGroup, h::UInt) = hash(g.axids, hash(g.kind, hash(:ChartGroup, h)))

"""
    ChartDataPoint

Per-point formatting override on a series (`c:dPt`), as read.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `idx` — the point's `c:idx`, 0-based as written.
- `invert_if_negative`, `bubble3d` — as written; `nothing` means absent.
- `raw` — the element as read, or `nothing` for a point built in code.

Identified by `series_idx` and `idx`. Functions taking `(c, d)` find the element
by these in the current part, so `d` stays usable across writes. The other fields
describe the point as it was when read; `raw` is never used to address the part.
"""
struct ChartDataPoint
    series_idx::Int                      # c:idx of the owning c:ser
    idx::Int                             # c:idx of the point, 0-based as written
    invert_if_negative::Union{Nothing,Bool}
    bubble3d::Union{Nothing,Bool}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartDataLabel

An individual data label override (`c:dLbl` within a series' `c:dLbls`), as read.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `idx` — the labelled point's `c:idx`, 0-based as written.
- `delete` — `c:delete`; a deleted label carries no other properties.
- `raw` — the element as read, or `nothing` for a label built in code.

Identified by `series_idx` and `idx`, as [`ChartDataPoint`](@ref).
"""
struct ChartDataLabel
    series_idx::Int
    idx::Int
    delete::Union{Nothing,Bool}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartTrendline

A trendline on a series (`c:trendline`), as read.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `ordinal` — 1-based position among the series' `c:trendline` elements.
- `kind` — `c:trendlineType`: `:linear`, `:exp`, `:log`, `:movingAvg`, `:poly`, `:power`.
- `name`, `order`, `period`, `forward`, `backward`, `intercept`, `disp_rsqr`,
  `disp_eq` — as written; `nothing` means absent.
- `raw` — the element as read, or `nothing` for a trendline built in code.

Identified by `series_idx` and `ordinal`. Trendlines have no identifier of their
own, so the ordinal is positional: it holds as long as no earlier trendline on
the same series is removed. `raw` is never used to address the part.
"""
struct ChartTrendline
    series_idx::Int
    ordinal::Int                         # 1-based among the series' c:trendline elements
    kind::Union{Nothing,Symbol}
    name::Union{Nothing,String}
    order::Union{Nothing,Int}
    period::Union{Nothing,Int}
    forward::Union{Nothing,Float64}
    backward::Union{Nothing,Float64}
    intercept::Union{Nothing,Float64}
    disp_rsqr::Union{Nothing,Bool}
    disp_eq::Union{Nothing,Bool}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartErrorBars

A set of error bars on a series (`c:errBars`), as read. A series carries at most
two, one per direction.

# Fields
- `series_idx` — `c:idx` of the owning series.
- `ordinal` — 1-based position among the series' `c:errBars` elements.
- `direction`, `bar_type`, `value_type`, `value`, `no_end_cap` — as written;
  `nothing` means absent.
- `raw` — the element as read, or `nothing` for error bars built in code.

Identified by `series_idx` and `ordinal`, as [`ChartTrendline`](@ref).
"""
struct ChartErrorBars
    series_idx::Int
    ordinal::Int                         # 1-based among the series' c:errBars elements
    direction::Union{Nothing,Symbol}
    bar_type::Union{Nothing,Symbol}
    value_type::Union{Nothing,Symbol}
    value::Union{Nothing,Float64}
    no_end_cap::Union{Nothing,Bool}
    raw::Union{Nothing,XML.Node}
end

"""
    ChartUpDownBars

Up-down bars on a line or stock chart group (`c:upDownBars`), as read.

# Fields
- `group` — the owning [`ChartGroup`](@ref), which is the key.
- `gap_width` — `c:gapWidth`; `nothing` means absent.
- `raw` — the element as read, or `nothing` for bars built in code.
"""
struct ChartUpDownBars
    group::ChartGroup
    gap_width::Union{Nothing,Int}
    raw::Union{Nothing,XML.Node}
end

"""
    FormatSite

One rung of a formatting cascade: the node that could carry a property, and
whether it does. `container` always exists — it is the `c:ser`, `c:dPt`,
`c:marker` or chart-space element being inspected. `props` is the `spPr` or
`txPr` found on it, or `nothing` where none was written, which is the rung
being absent rather than the property being off.

`kind` distinguishes what `props` holds, so a chain is interpretable without
knowing which resolver built it: `:shape` for an `spPr` on the element itself,
`:marker` for an `spPr` on its `c:marker`, `:text` for a `txPr`.
"""
struct FormatSite
    level::Symbol                      # :point, :series, :group, :plotarea, :chartspace, :style, :theme
    kind::Symbol                       # :shape, :marker, :text
    container::XML.Node
    props::Union{Nothing,XML.Node}
end

"""
    Effective{T}

The result of resolving a property up a cascade. `value` is the first explicit
setting found, or `nothing` where the property is written at no rung at all —
which means Excel takes it from the chart style part, not that it is off. An
explicit `<a:noFill/>` resolves to a `DrawingFill` with `kind === :none` and a
`site`, because deliberately off is a setting.

`site` is the rung that answered. `chain` is every rung that was or could have
been consulted, highest precedence first, and is populated whether or not a
value was found — a setter uses it to decide where to write.
"""
struct Effective{T}
    value::Union{Nothing,T}
    site::Union{Nothing,FormatSite}
    chain::Vector{FormatSite}
end

"""
    SchemaKey

The `(namespace, complex-type)` pair identifying which `xsd:sequence` governs an
element's children. Not derivable from the element's tag: every series is `c:ser`
but its child order depends on the group containing it (`barSer`, `lineSer`, …),
and `c:spPr` is `a:CT_ShapeProperties` despite its chart prefix. Callers state it.
"""
const SchemaKey = Tuple{String,String}

