const highlights::Dict{String,Dict{String,Dict{String,String}}} = Dict(
    "redfilltext" => Dict(
        "font" => Dict("color" => "FF9C0006"),
        "fill" => Dict("pattern" => "solid", "bgColor" => "FFFFC7CE")
    ),
    "yellowfilltext" => Dict(
        "font" => Dict("color" => "FF9C5700"),
        "fill" => Dict("pattern" => "solid", "bgColor" => "FFFFEB9C")
    ),
    "greenfilltext" => Dict(
        "font" => Dict("color" => "FF006100"),
        "fill" => Dict("pattern" => "solid", "bgColor" => "FFC6EFCE")
    ),
    "redfill" => Dict(
        "fill" => Dict("pattern" => "solid", "bgColor" => "FFFFC7CE")
    ),
    "redtext" => Dict(
        "font" => Dict("color" => "FF9C0006"),
    ),
    "redborder" => Dict(
        "border" => Dict("color" => "FF9C0006", "style" => "thin")
    )
)
const databars::Dict{String,Dict{String,String}} = Dict(
    "bluegrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF638EC6",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FF638EC6",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "greengrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF63C384",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FF63C384",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "redgrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFFF555A",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FFFF555A",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "orangegrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFFFB628",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FFFFB628",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "lightbluegrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF008AEF",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FF008AEF",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "purplegrad" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFD6007B",
        "borders" => "true",
        "sameNegBorders" => "false",
        "border_col" => "FFD6007B",
        "neg_fill_col" => "FFFF0000",
        "neg_border_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "blue" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF638EC6",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "green" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF63C384",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "red" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFFF555A",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "orange" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFFFB628",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "lightblue" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FF008AEF",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
    "purple" => Dict(
        "min_type" => "automatic",
        "max_type" => "automatic",
        "fill_col" => "FFD6007B",
        "gradient" => "false",
        "neg_fill_col" => "FFFF0000",
        "axis_col" => "FF000000"
    ),
)
const colorscales::Dict{String,XML.Node} = Dict(    # Defines the 12 standard, built-in Excel color scales for conditional formatting.
    "greenyellowred" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFFEB84"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "redyellowgreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFFEB84"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "greenwhitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "redwhitegreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "bluewhitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF5A8AC6")
        )
    ),
    "redwhiteblue" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF5A8AC6"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "whitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF")
        )
    ),
    "redwhite" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "greenwhite" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "whitegreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFCFCFF")
        )
    ),
    "greenyellow" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFFEF9C"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "yellowgreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFFEF9C")
        )
    )
)
const iconsets::Dict{String,XML.Node} = Dict(    # Defines the 20 standard, built-in Excel icon sets for conditional formatting.
    "3Arrows" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Arrows",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "4Arrows" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4Arrows",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "5Arrows" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="5Arrows",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="20"),
            XML.h.cfvo(type="percent", val="40"),
            XML.h.cfvo(type="percent", val="60"),
            XML.h.cfvo(type="percent", val="80"),
        )
    ),
    "3ArrowsGray" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3ArrowsGray",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "4ArrowsGray" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4ArrowsGray",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "5ArrowsGray" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="5ArrowsGray",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="20"),
            XML.h.cfvo(type="percent", val="40"),
            XML.h.cfvo(type="percent", val="60"),
            XML.h.cfvo(type="percent", val="80"),
        )
    ),
    "3TrafficLights" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3TrafficLights",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "3Signs" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Signs",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "4RedToBlack" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4RedToBlack",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "3TrafficLights2" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3TrafficLights2",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "4TrafficLights" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4TraficLights",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "3Symbols" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Symbols",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "3Symbols2" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Symbols2",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "3Flags" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="3Flags",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="33"),
            XML.h.cfvo(type="percent", val="67"),
        )
    ),
    "5Quarters" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="5Quarters",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="20"),
            XML.h.cfvo(type="percent", val="40"),
            XML.h.cfvo(type="percent", val="60"),
            XML.h.cfvo(type="percent", val="80"),
        )
    ),
    "4Rating" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="4Rating",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="25"),
            XML.h.cfvo(type="percent", val="50"),
            XML.h.cfvo(type="percent", val="75"),
        )
    ),
    "5Rating" => XML.h.cfRule(type="iconSet", priority="1",
        XML.h.iconSet(iconSet="5Rating",
            XML.h.cfvo(type="percent", val="0"),
            XML.h.cfvo(type="percent", val="20"),
            XML.h.cfvo(type="percent", val="40"),
            XML.h.cfvo(type="percent", val="60"),
            XML.h.cfvo(type="percent", val="80"),
        )
    ),
    # These three require Excel 2010 extensions and will be ignored by earlier versions of Excel.
    "3Triangles" => get_x14_icon("3Triangles"),
    "3Stars" => get_x14_icon("3Stars"),
    "5Boxes" => get_x14_icon("5Boxes"),
    "Custom" => get_x14_icon("Custom")
)
const allIcons::Dict{String,Tuple{String,String}} = Dict(
    "1" => ("3Arrows", "0"),
    "2" => ("3Arrows", "1"),
    "3" => ("3Arrows", "2"),
    "4" => ("4Arrows", "1"),
    "5" => ("4Arrows", "2"),
    "6" => ("3ArrowsGray", "0"),
    "7" => ("3ArrowsGray", "1"),
    "8" => ("3ArrowsGray", "2"),
    "9" => ("4ArrowsGray", "1"),
    "10" => ("4ArrowsGray", "2"),
    "11" => ("3Flags", "0"),
    "12" => ("3Flags", "1"),
    "13" => ("3Flags", "2"),
    "14" => ("3TrafficLights1", "0"),
    "15" => ("3TrafficLights1", "1"),
    "16" => ("3TrafficLights1", "2"),
    "17" => ("3TrafficLights2", "0"),
    "18" => ("3TrafficLights2", "1"),
    "19" => ("3TrafficLights2", "2"),
    "20" => ("4TrafficLights", "0"),
    "21" => ("3Signs", "0"),
    "22" => ("3Signs", "1"),
    "23" => ("3Symbols", "0"),
    "24" => ("3Symbols", "1"),
    "25" => ("3Symbols", "2"),
    "26" => ("3Symbols2", "0"),
    "27" => ("3Symbols2", "1"),
    "28" => ("3Symbols2", "2"),
    "29" => ("4RedToBlack", "0"),
    "30" => ("4RedToBlack", "1"),
    "31" => ("4RedToBlack", "2"),
    "32" => ("4RedToBlack", "3"),
    "33" => ("5Quarters", "0"),
    "34" => ("5Quarters", "1"),
    "35" => ("5Quarters", "2"),
    "36" => ("5Quarters", "3"),
    "37" => ("5Rating", "0"),
    "38" => ("5Rating", "1"),
    "39" => ("5Rating", "2"),
    "40" => ("5Rating", "3"),
    "41" => ("5Rating", "4"),
    "42" => ("3Stars", "0"),
    "43" => ("3Stars", "1"),
    "44" => ("3Stars", "2"),
    "45" => ("3Triangles", "0"),
    "46" => ("3Triangles", "1"),
    "47" => ("3Triangles", "2"),
    "48" => ("5Boxes", "0"),
    "49" => ("5Boxes", "1"),
    "50" => ("5Boxes", "2"),
    "51" => ("5Boxes", "3"),
    "52" => ("5Boxes", "4")
)

const timeperiods::Dict{String,String} = Dict(
    "last7Days" => "AND(TODAY()-FLOOR(__CR__,1)<=6,FLOOR(__CR__,1)<=TODAY())",
    "yesterday" => "FLOOR(__CR__,1)=TODAY()-1",
    "today" => "FLOOR(__CR__,1)=TODAY()",
    "tomorrow" => "FLOOR(__CR__,1)=TODAY()+1",
    "lastWeek" => "AND(TODAY()-ROUNDDOWN(__CR__,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(__CR__,0)<(WEEKDAY(TODAY())+7))",
    "thisWeek" => "AND(TODAY()-ROUNDDOWN(__CR__,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(__CR__,0)-TODAY()<=7-WEEKDAY(TODAY()))",
    "nextWeek" => "AND(ROUNDDOWN(__CR__,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(__CR__,0)-TODAY()<(15-WEEKDAY(TODAY())))",
    "lastMonth" => "AND(MONTH(__CR__)=MONTH(EDATE(TODAY(),0-1)),YEAR(__CR__)=YEAR(EDATE(TODAY(),0-1)))",
    "thisMonth" => "AND(MONTH(__CR__)=MONTH(TODAY()),YEAR(__CR__)=YEAR(TODAY()))",
    "nextMonth" => "AND(MONTH(__CR__)=MONTH(EDATE(TODAY(),0+1)),YEAR(__CR__)=YEAR(EDATE(TODAY(),0+1)))"
)

const CfKey = Union{CellRange,NonContiguousRange}

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
            @assert localname(cf) == "conditionalFormatting" "Something wrong here"
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

"""
    setDataBands(sheet, rng; breaks, bands, [type], [scale], [compress], kwargs...)

Split `rng` into value-based bands and apply a separate conditional format to each,
so that a single contiguous range can show (for example) data bars whose colour
changes with value as well as their length.

Excel has no single conditional format that varies colour by value: a `dataBar` rule
carries one colour and applies to every cell in its range. This function works around
that by partitioning `rng` on `breaks` and writing one rule per band, each over the
non-contiguous range of cells falling in that band.

`breaks` is a sorted vector of `n` boundaries giving `n+1` bands, each boundary
belonging to the band below it, as in [`partition`](@ref). `bands` is a vector of
`n+1` NamedTuples holding whatever differs between bands. Any further keyword
arguments are applied to every band; a band may override one by setting it itself.

`scale` fixes the value axis shared by the bands, and applies to `type` values that
take conditional format value objects (`:databar`, `:colorscale`, `:iconset`):

  - `:global` (default) -- every band uses the minimum and maximum of the whole of
    `rng`, so bar lengths remain comparable across bands
  - `:band` -- each band scales to its own values, which normally means every band
    spans the full bar width and the lengths become meaningless
  - `(lo, hi)` -- explicit bounds, useful for a fixed axis across several sheets

Returns the label => range pairs written, as [`partition`](@ref) does.

The bands reflect cell values at the time of writing. If a value is later edited in
Excel the bar length updates but its colour does not, because the cell remains in the
band it was assigned to.

# Examples

```julia
julia> XLSX.setDataBands(sheet, "B2:B101";
           breaks = [10, 50],
           bands  = [(colour="FF63BE7B",), (colour="FFFFC000",), (colour="FFFF0000",)])
```

A fixed axis, with solid fill on every band:

```julia
julia> XLSX.setDataBands(sheet, "B2:B101";
           breaks    = [10, 50],
           bands     = [(colour="FF63BE7B",), (colour="FFFFC000",), (colour="FFFF0000",)],
           scale     = (0, 100),
           gradient  = false)
```

See also [`partition`](@ref), [`setConditionalFormat`](@ref).
"""
function setDataBands(ws::Worksheet, rng::CellRange;
                      breaks::AbstractVector{<:Real},
                      bands::AbstractVector{<:NamedTuple},
                      type::Symbol = :databar,
                      scale = :global,
                      compress::Bool = true,
                      kwargs...)

    length(bands) == length(breaks) + 1 ||
        throw(XLSXError("Need $(length(breaks)+1) bands for $(length(breaks)) breaks, got $(length(bands))."))

    labs = collect(1:length(bands))
    p = partition(ws, rng, breaks; labels = labs, compress)
    isempty(p) && return p

    scaled = type in (:databar, :colorscale, :iconset)
    common = if scaled && scale !== :band
        lo, hi = if scale === :global
            vals = filter(v -> v isa Real, vec(getdata(ws, rng)))
            isempty(vals) && throw(XLSXError("No numeric values in `$rng` to scale against."))
            minimum(vals), maximum(vals)
        else
            first(scale), last(scale)
        end
        (; min_type = :num, min_value = string(lo),
           max_type = :num, max_value = string(hi), kwargs...)
    else
        (; kwargs...)
    end

    for (label, ncr) in p
        setConditionalFormat(ws, ncr; type, common..., bands[label]...)
    end
    return p
end

setDataBands(ws::Worksheet, rng::AbstractString; kwargs...) =
    setDataBands(ws, CellRange(rng); kwargs...)

"""
    setConditionalFormat(ws::Worksheet, cr::String, type::Symbol; kw...) -> ::Int
    setConditionalFormat(xf::XLSXFile,  cr::String, type::Symbol; kw...) -> ::Int

    setConditionalFormat(ws::Worksheet, rows, cols, type::Symbol; kw...) -> ::Int

Add a new conditional format to a cell range, row range or column range in a 
worksheet or `XLSXFile`.  Alternatively, ranges can be specified by giving rows 
and columns separately.

There are many options for applying differnt types of custom format. For a basic guide, 
refer to the section on [Conditional formats](@ref) in the Formatting Guide.

The `type` argument specifies which of Excel's conditional format types will be applied.

Valid options for `type` are:
- `:cellIs`
- `:top10`
- `:aboveAverage`
- `:containsText`
- `:notContainsText`
- `:beginsWith`
- `:endsWith`
- `:timePeriod`
- `:containsErrors`
- `:notContainsErrors`
- `:containsBlanks`
- `:notContainsBlanks`
- `:uniqueValues`
- `:duplicateValues`
- `:dataBar`
- `:colorScale`
- `:iconSet`

Keyword options differ according to the `type` specified, as set out below.

# type = :cellIs

Defines a conditional format based on the value of each cell in a range.

Valid keywords are:
- `operator`   : Defines the comparison to make.
- `value`      : defines the first value to compare against. This can be a cell reference (e.g. `"A1"`) or a number.
- `value2`     : defines the second value to compare against. This can be a cell reference (e.g. `"A1"`) or a number.
- `stopIfTrue` : Stops evaluating further conditional formats in any cell if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

All keywords are defined using Strings (e.g. `value = "2"` or `value = "A2"`).

The keyword `operator` defines the comparison to use in the conditional formatting. 
If the condition is met, the format is applied.
Valid options are:

- `greaterThan`     (cell >  `value`) (default)
- `greaterEqual`    (cell >= `value`)
- `lessThan`        (cell <  `value`)
- `lessEqual`       (cell <= `value`)
- `equal`           (cell == `value`)
- `notEqual`        (cell != `value`)
- `between`         (cell between `value` and `value2`)
- `notBetween`      (cell not between `value` and `value2`)

If not specified (when required), `value` will be the arithmetic average of the 
(non-missing) cell values in the range if values are numeric. If the cell values 
are non-numeric, an error is thrown.

Formatting to be applied if the condition is met can be defined in one of two ways. 
Use the keyword `dxStyle` to select one of the built-in Excel formats. 
Valid options are:

- `redfilltext`     (light red fill, dark red text) (default)
- `yellowfilltext`  (light yellow fill, dark yellow text)
- `greenfilltext`   (light green fill, dark green text)
- `redfill`         (light red fill)
- `redtext`         (dark red text)
- `redborder`       (dark red cell borders)

Alternatively, you can define a custom format by using the keywords `format`, `font`,
`border`, and `fill` which each take a vector of pairs of strings. The first string 
is the name of the attribute to set and the second is the value to set it to.  
Valid attributes for each keyword are:

- `format` : `format`
- `font`   : `color`, `bold`, `italic`, `under`, `strike`
- `fill`   : `pattern`, `bgColor`, `fgColor`
- `border` : `style`, `color`

Refer to [`setFormat()`](@ref), [`setFont()`](@ref), [`setFill()`](@ref) and [`setBorder()`](@ref) for
more details on the valid attributes and values.

!!! note

    Excel limits the formatting attributes that can be set in a conditional format.
    It is not possible to set the size or name of a font and neither is it possible to set 
    any of the cell alignment attributes. Diagonal borders cannot be set either.

    Although it is not a limitation of Excel, this function sets all the border attributes 
    for each side of a cell to be the same.

If both `dxStyle` and custom formatting keywords are specified, `dxStyle` will be used 
and the custom formatting will be ignored.
If neither `dxStyle` nor custom formatting keywords are specified, the default 
is `dxStyle="redfilltext"`.

# Examples

```julia
julia> XLSX.setConditionalFormat(s, "B1:B5", :cellIs) # Defaults to `operator="greaterThan"`, `dxStyle="redfilltext"` and `value` set to the arithmetic agverage of cell values in `rng`.

julia> XLSX.setConditionalFormat(s, "B1:B5", :cellIs;
            operator="between",
            value="2",
            value2="3",
            fill = ["pattern" => "none", "bgColor"=>"FFFFC7CE"],
            format = ["format"=>"0.00%"],
            font = ["color"=>"blue", "bold"=>"true"]
        )

julia> XLSX.setConditionalFormat(s, "B1:B5", :cellIs; 
            operator="greaterThan",
            value="4",
            fill = ["pattern" => "none", "bgColor"=>"green"],
            format = ["format"=>"0.0"],
            font = ["color"=>"red", "italic"=>"true"]
        )

julia> XLSX.setConditionalFormat(s, "B1:B5", :cellIs;
            operator="lessThan",
            value="2",
            fill = ["pattern" => "none", "bgColor"=>"yellow"],
            format = ["format"=>"0.0"],
            font = ["color"=>"green"],
            border = ["style"=>"thick", "color"=>"coral"]
        )

```

# type = :top10

This conditional format can be used to highlight cells in the top (bottom) n within the 
range or in the top (bottom) n% (ie in the top 5 or in the top 5% of values in the range). 

The available keywords are:

- `operator`   : Defines the comparison to make.
- `value`      : Gives the for comparison or a cell reference (e.g. `"A1"`).
- `stopIfTrue` : Stops evaluating further conditional formats in any cell if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply.
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

Valid values for the `operator` keyword are the following:

- `topN`            (cell is in the top n (= `value`) values of the range)
- `bottomN`         (cell is in the bottom n (= `value`) values of the range)
- `topN%`           (cell is in the top n% (= `value`) values of the range)
- `bottomN%`        (cell is in the bottom n% (= `value`) values of the range)

Default keyowrds are `operator="TopN"` and `value="10"`.
    
Multiple conditional formats may be applied to the same or overlapping cell ranges. 
If `stopIfTrue=true` the first condition that is met will be applied but all subsequent 
conditional formats for that cell will be skipped. If `stopIfTrue=false` (default) all 
relevant conditional formats will be applied to the cell in turn.

For example usage of the `stopIfTrue` keyword, refer to [Overlaying conditional formats](@ref) 
in the Formatting Guide.

The remaining keywords are defined as above for `type = :cellIs`.

# Examples

```julia
julia> f=XLSX.newxlsx()
XLSXFile("C:\\...\\blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 1x1           A1:A1


julia> s=f[1]
1×1 XLSX.Worksheet: ["Sheet1"](A1:A1)

julia> for i=1:10;for j=1:10; s[i,j]=i*j;end;end

julia> s[:]
10×10 Matrix{Any}:
  1   2   3   4   5   6   7   8   9   10
  2   4   6   8  10  12  14  16  18   20
  3   6   9  12  15  18  21  24  27   30
  4   8  12  16  20  24  28  32  36   40
  5  10  15  20  25  30  35  40  45   50
  6  12  18  24  30  36  42  48  54   60
  7  14  21  28  35  42  49  56  63   70
  8  16  24  32  40  48  56  64  72   80
  9  18  27  36  45  54  63  72  81   90
 10  20  30  40  50  60  70  80  90  100

julia> XLSX.setConditionalFormat(s, "A1:J10", :top10; operator="bottomN", value="1", stopIfTrue="true", dxStyle="redfilltext")
0

julia> XLSX.setConditionalFormat(s, "A1:J10", :top10; operator="topN", value="1", stopIfTrue="true", dxStyle="greenfilltext")
0

julia> XLSX.setConditionalFormat(s, "A1:J10", :top10;
                operator="topN%",
                value="20",
                fill=["pattern"=>"solid", "bgColor"=>"cyan"])
0

julia> XLSX.setConditionalFormat(s, "A1:J10", :top10;
                operator="bottomN%",
                value="20",
                fill=["pattern"=>"solid", "bgColor"=>"yellow"])
0

```
![image|320x500](../images/topN.png)

# type = :aboveAverage

This conditional format can be used to compare cell values in the range with the 
average value for the range. 

The available keywords are:

- `operator`   : Defines the comparison to make.
- `stopIfTrue` : Stops evaluating further conditional formats in any cell if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply.
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

Valid values for the `operator` keyword are the following:

- `aboveAverage`    (cell is above the average of the range) (default)
- `aboveEqAverage`  (cell is above or equal to the average of the range)
- `plus1StdDev`     (cell is above the average of the range + 1 standard deviation)
- `plus2StdDev`     (cell is above the average of the range + 2 standard deviations)
- `plus3StdDev`     (cell is above the average of the range + 3 standard deviations)
- `belowAverage`    (cell is below the average of the range)
- `belowEqAverage`  (cell is below or equal to the average of the range)
- `minus1StdDev`    (cell is below the average of the range - 1 standard deviation)
- `minus2StdDev`    (cell is below the average of the range - 2 standard deviations)
- `minus3StdDev`    (cell is below the average of the range - 3 standard deviations)

The remaining keywords are defined as above for `type = :cellIs`.

# Examples

```julia
julia> using Random, Distributions

julia> d=Normal()
Normal{Float64}(μ=0.0, σ=1.0)

julia> columns=rand(d,1000)                                                                                                                                                        
1000-element Vector{Float64}:
-1.5515478694605092
  0.36859583733587165
  1.5349535865662158
 -0.2352610551087202
  0.12355875388105911
  0.5859222303845908
 -0.6326662651426166
  1.0610118292961683
 -0.7891578831398097
  0.031022172414689787
 -0.5534440118018843
 -2.3538883599955023
  ⋮
  0.4813001892130465
  0.03871017417416217
  0.7224728281160403
 -1.1265372949908539
  1.5714393857211955
  0.31438739499933255
  0.4852591013082452
  0.5363388236349432
  1.1268430910133729
  0.7691442442244849
  1.0061732938516454

julia> f=XLSX.newxlsx()
XLSXFile("C:\\...\\blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
               Sheet1 1x1           A1:A1


julia> s=f[1]
1×1 XLSX.Worksheet: ["Sheet1"](A1:A1) 

julia> XLSX.writetable!(s, [columns], ["normal"])

julia> XLSX.setConditionalFormat(s, "A2:A1001", :aboveAverage ;
                operator="plus3StdDev",
                stopIfTrue="true",
                fill = ["pattern"=>"solid", "bgColor"=>"red"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A2:A1001", :aboveAverage ;
                operator="minus3StdDev",
                stopIfTrue="true",
                fill = ["pattern"=>"solid", "bgColor"=>"red"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A2:A1001", :aboveAverage ;
                operator="plus2StdDev",
                stopIfTrue="true",
                fill = ["pattern"=>"solid", "bgColor"=>"tomato"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A2:A1001", :aboveAverage ;
                operator="minus2StdDev",
                stopIfTrue="true",
                fill = ["pattern"=>"solid", "bgColor"=>"tomato"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A2:A1001", :aboveAverage ;
                operator="minus1StdDev",
                stopIfTrue="true",
                fill = ["pattern"=>"solid", "bgColor"=>"pink"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A2:A1001", :aboveAverage ;
                operator="plus1StdDev",
                stopIfTrue = "true",
                fill = ["pattern"=>"solid", "bgColor"=>"pink"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A2:A1001", :aboveAverage ;
                operator="belowEqAverage",
                fill = ["pattern"=>"solid", "bgColor"=>"green"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A2:A1001", :aboveAverage ;
                operator="aboveEqAverage", 
                fill = ["pattern"=>"solid", "bgColor"=>"green"],
                font = ["color"=>"white", "bold"=>"true"])
0

```

# type = :containsText, :notContainsText, :beginsWith or :endsWith

Highlight cells in the range that contain (or do not contain), begin or end with 
a specific text string. The default is `containsText`.

Valid keywords are:

- `value`      : Gives the literal text to match or provides a cell reference (e.g. `"A1"`).
- `stopIfTrue` : Stops evaluating further conditional formats in any cell if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply.
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

The keyword `value` gives the literal text to compare (eg. "Hello World") or provides a cell reference 
(e.g. `"A1"`). It is a required keyword with no default value.

The remaining keywords are optional and are defined as above for `type = :cellIs`.

# Examples

```julia
julia> s[:]
4×1 Matrix{Any}:
 "Hello World"
 "Life the universe and everything"
 "Once upon a time"
 "In America"

julia> XLSX.setConditionalFormat(s, "A1:A4", :containsText;
                value="th",
                fill = ["pattern"=>"solid", "bgColor"=>"cyan"],
                font = ["color"=>"black", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A1:A4", :notContainsText;
                value="i",
                fill = ["pattern"=>"solid", "bgColor"=>"green"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A1:A4", :beginsWith ;
                value="On",
                fill = ["pattern"=>"solid", "bgColor"=>"red"],
                font = ["color"=>"white", "bold"=>"true"])
0

julia> XLSX.setConditionalFormat(s, "A1:A4", :endsWith ;
                value="ica",
                fill = ["pattern"=>"solid", "bgColor"=>"blue"],
                font = ["color"=>"white", "bold"=>"true"])
0

```
![image|320x500](../images/containsText.png)

# type = :timePeriod

When cells contain dates, this conditional format can be used to highlight cells.
The available keywords are:

- `operator`   : Defines the comparison to make.
- `stopIfTrue` : Stops evaluating further conditional formats in any cell if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

Valid values for the keyword `operator` are the following:

- `yesterday`
- `today`
- `tomorrow`
- `last7Days` (default)
- `lastWeek`
- `thisWeek`
- `nextWeek`
- `lastMonth`
- `thisMonth`
- `nextMonth`

The remaining keywords are defined as above for `type = :cellIs`.

# Examples

```julia
julia> s[1:13, 1]
13×1 Matrix{Any}:
 "Dates"
 2024-11-20
 2024-12-20
 2025-01-08
 2025-02-08
 2025-03-08
 2025-04-08
 2025-05-08
 2025-05-09
 2025-05-10
 2025-05-14
 2025-06-08
 2025-07-08

julia> XLSX.setConditionalFormat(s, "A1:A13", :timePeriod; operator="today", dxStyle = "greenfilltext")
0

julia> XLSX.setConditionalFormat(s, "A1:A13", :timePeriod; operator="tomorrow", dxStyle = "yellowfilltext")
0

julia> XLSX.setConditionalFormat(s, "A1:A13", :timePeriod; operator="nextMonth", dxStyle = "redfilltext")
0

julia> XLSX.setConditionalFormat(s, "A1:A13", :timePeriod;
                operator="lastMonth", 
                fill = ["pattern"=>"solid", "bgColor"=>"blue"], 
                font = ["color"=>"yellow", "bold"=>"true"])        
0

```
![image|320x500](../images/timePeriod-9thMay2025.png)


# type = :containsErrors, :notContainsErrors, :containsBlanks, :notContainsBlanks, :uniqueValues or :duplicateValues

These conditional formatting options highlight cells that contain or don't contain errors, 
are blank (default) or not blank, are unique in the range or are duplicates within the range. 
The available keywords are: 

- `stopIfTrue` : Stops evaluating further conditional formats in any cell if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

These keywords are defined as above for the `:cellIs` conditional format type.

# Examples

```julia
julia> XLSX.setConditionalFormat(s, "A1:A7", :containsErrors;
                stopIfTrue="true",
                fill = ["pattern"=>"solid", "bgColor"=>"blue"],
                font = ["color"=>"white", "bold"=>"true"])        
0

julia> XLSX.setConditionalFormat(s, "A1:A7", :containsBlanks;
                stopIfTrue="true",
                fill = ["pattern"=>"solid", "bgColor"=>"green"],
                font = ["color"=>"black", "bold"=>"true"])       
0

julia> XLSX.setConditionalFormat(s, "A1:A7", :uniqueValues;
                stopIfTrue="true",
                fill = ["pattern"=>"solid", "bgColor"=>"yellow"],
                font = ["color"=>"black", "bold"=>"true"])        
0

julia> XLSX.setConditionalFormat(s, "A1:A7", :duplicateValues;
                fill = ["pattern"=>"solid", "bgColor"=>"cyan"],
                font = ["color"=>"black", "bold"=>"true"])
0

```
![image|320x500](../images/errorBlank.png)

# type = :expressiom

Set a conditional format when an expression evaluated in each cell is `true`.

The available keywords are:

- `formula`    : Specifies the formula to use. This must be a valid Excel formula.
- `stopIfTrue` : Stops evaluating further conditional formats in any cell if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

The keyword `formula` is required and there is no default value. Formulae must be valid 
Excel formulae and written in US english with comma separators. Cell references may be 
absolute or relative references in either the row or the column or both.

The remaining keywords are defined as above for `type = :cellIs`.

# Examples

```julia
julia> XLSX.setConditionalFormat(s, "A1:C4", :expression; formula = "A1 < 16", dxStyle="greenfilltext")

julia> XLSX.setConditionalFormat(s, 1:5, 1:4, :expression;
            formula="A1=1",
            fill=["pattern" => "none", "bgColor" => "yellow"],
            format=["format" => "0.0"],
            font=["color" => "green"],
            border=["style" => "thick", "color" => "coral"]
        )

julia> XLSX.setConditionalFormat(s, "B2:D11", :expression; formula = "average(B\$2:B\$11) > average(A\$2:A\$11)", dxStyle = "greenfilltext")

julia> XLSX.setConditionalFormat(s, "A1:E5", :expression; formula = "E5<50", dxStyle = "redfilltext")

```

# type = :dataBar

Apply data bars to cells in a range depending on their values. The keyword `databar`
can be used to select one of 12 built-in databars Excel provides by name. Valid names are:
- `bluegrad` (default)
- `greengrad`
- `redgrad`
- `orangegrad`
- `lightbluegrad`
- `purplegrad`
- `blue`
- `green`
- `red`
- `orange`
- `lightblue`
- `purple`

The first six (with a `grad` suffix) yield bars with a color gradient while the remainder 
yield bars of solid color. By default, all built in data bars define their range from the 
minumum and maximum values in the range and negative values are given a red bar. These default
settings can each be modified using the other keyword options available.

Remaining keyword options provided are:
- `showVal` - set to "false" to show databars only and hide cell values
- `gradient` - set to "false" to use a solid color bar rather than a gradient fill
- `borders` - set to "true" to show borders around each bar
- `sameNegFill` - set to "true" to use the same fill color on negative bars as positive.
- `sameNegBorders` - set to "false" to use the same border color on negative bars as positive
- `direction` - determines the direction of the bars from the axis, "leftToRight" or "rightToLeft"
- `min_type` - Defines how the minimum of the bar scale is defined ("num", "min", "percent", percentile", "formula" or "automatic")
- `min_val` - Defines the minimum value for the data bar scale. May be a number(as a string), a cell reference or a formula (if type="formula").
- `max_type` - Defines how the maximum of the bar scale is defined ("num", "max", "percent", percentile", "formula" or "automatic")
- `max_val` - Defines the maximum value for the data bar scale. May be a number(as a string), a cell reference or a formula (if type="formula").
- `fill_col` - Defines the color of the fill for positive bars (8 digit hex or by name)
- `border_col` - Defines the color of the border for positive bars (8 digit hex or by name)
- `neg_fill_col` - Defines the color of the fill for negative bars (8 digit hex or by name)
- `neg_border_col` - Defines the color of the border for negative bars (8 digit hex or by name)
- `axis_pos` - Defines the position of the axis ("middle" or "none")
- `axis_col` - Defines the color of the axis (8 digit hex or by name)

# Examples
```julia
julia> XLSX.setConditionalFormat(s, "A1:A11", :dataBar)

julia> XLSX.setConditionalFormat(s, "B1:B11", :dataBar; databar="purple")

julia> XLSX.setConditionalFormat(s, "D1:D11", :dataBar; 
            gradient="true", 
            direction="rightToLeft", 
            axis_pos="none", 
            showVal="false"
        )

julia> XLSX.setConditionalFormat(s, "F1:F11", :dataBar;
            gradient="false",
            sameNegFill="true",
            sameNegBorders="true"
        )

julia> XLSX.setConditionalFormat(f, "Sheet1!G1:G11", :dataBar;
            fill_col="coral", border_col = "cyan",
            neg_fill_col="cyan", neg_border_col = "coral"
        )

julia> XLSX.setConditionalFormat(f, "Sheet1!J1:J11", :dataBar; axis_col="magenta")

julia> XLSX.setConditionalFormat(s, 15:25, 1, :dataBar;
            min_type="least", max_type="highest"
        )

julia> XLSX.setConditionalFormat(s, 15:25, 2, :dataBar; 
            databar="purple", 
            min_type="percent", max_type="percent",
            min_val="20", max_val="60"
        )

julia> XLSX.setConditionalFormat(s, "C15:C25", :dataBar;
            databar="blue",
            min_type="num", max_type="num",
            min_val="-1", max_val="6",
            gradient="true",
            direction="leftToRight", 
            axis_pos="none"
        )

julia> XLSX.setConditionalFormat(s, "E15:E25", :dataBar;
            gradient="true",
            min_type="percentile", max_type="percentile",
            min_val="20", max_val="80",
            direction="rightToLeft",
            axis_pos="middle"
        )

julia> XLSX.setConditionalFormat(s, "G15:G25", :dataBar; 
            min_type="num", max_type="formula", 
            min_val="\$L\$1", max_val="\$M\$1 * \$N\$1 + 3",
            fill_col="coral", border_col = "cyan",
            neg_fill_col="cyan", neg_border_col = "coral"
        )

```

# type = :colorScale

Define a 2-color or 3-color colorscale conditional format.

Use the keyword `colorscale` to choose one of the 12 built-in Excel colorscales:

- `"redyellowgreen"`: Red, Yellow, Green 3-color scale.
- `"greenyellowred"`: Green, Yellow, Red 3-color scale.
- `"redwhitegreen"` : Red, White, Green 3-color scale.
- `"greenwhitered"` : Green, White, Red 3-color scale.
- `"redwhiteblue"`  : Red, White, Blue 3-color scale.
- `"bluewhitered"`  : Blue, White, Red 3-color scale.
- `"redwhite"`      : Red, White 2-color scale.
- `"whitered"`      : White, Red 2-color scale.
- `"whitegreen"`    : White, Green 2-color scale.
- `"greenwhite"`    : Green, White 2-color scale.
- `"yellowgreen"`   : Yellow, Green 2-color scale.
- `"greenyellow"`   : Green, Yellow 2-color scale (default).

Alternatively, you can define a custom color scale by omitting the `colorscale` keyword and 
instead using the following keywords:

- `min_type`: The type of the minimum value. Valid values are: `min`, `percentile`, `percent`, `num` or `formula`.
- `min_val` : The value of the minimum. Omit if `min_type="min"`.
- `min_col` : The color of the minimum value.
- `mid_type`: Valid values are: `percentile`, `percent`, `num` or `formula`. Omit for a 2-color scale.
- `mid_val` : The value of the scale mid point. Omit for a 2-color scale.
- `mid_col` : The color of the mid point. Omit for a 2-color scale.
- `max_type`: The type of the maximum value. Valid values are: `max`, `percentile`, `percent`, `num` or `formula`.
- `max_val` : The value of the maximum value. Omit if `max_type="max"`.
- `max_col` : The color of the maximum value.

The keywords `min_val`, `mid_val`, and `max_val` can be a number or cell reference (e.g. `"\$A\$1"`) for any value 
of the related type keyword or, if the related type keyword is set to `formula`, may be a valid Excel formula that 
calculates a number. Cell references used in a formula must be specified as absolute references.

Colors can be specified using an 8-digit hex string (e.g. `FF0000FF` for blue) or any named 
color from [Colors.jl](https://juliagraphics.github.io/Colors.jl/stable/namedcolors/).

# Examples

```julia
julia> XLSX.setConditionalFormat(f["Sheet1"], "A1:F12", :colorScale) # Defaults to the `greenyellow` built-in scale.
0

julia> XLSX.setConditionalFormat(f["Sheet1"], "A13:C18", :colorScale; colorscale="whitered")
0

julia> XLSX.setConditionalFormat(f["Sheet1"], "D13:F18", :colorScale; colorscale="bluewhitered")
0

julia> XLSX.setConditionalFormat(f["Sheet1"], "A13:F22", :colorScale;
            min_type="num", 
            min_val="2",
            min_col="tomato",
            mid_type="num",
            mid_val="6", 
            mid_col="lawngreen",
            max_type="num",
            max_val="10",
            max_col="cadetblue"
        )
0

```

# type = :iconSet

Apply a set of icons to cells in a range depending on their values. The keyword `iconset`
can be used to select one of 20 built-in icon sets Excel provides by name. Valid names are:
- `3Arrows`
- `5ArrowsGray`
- `3TrafficLights` (default)
- `3Flags`
- `5Quarters`
- `4Rating`
- `5Rating`
- `3Symbols`
- `3Symbols2`
- `3Signs`
- `3TrafficLights2`
- `4TrafficLights`
- `4BlackToRed`
- `4Arrows`
- `5Arrows`
- `3ArrowsGray`
- `4ArrowsGray`
- `3Triangles`
- `3Stars`
- `5Boxes`

The digit prefix to the name indicates how many icons there are in a set, and therefore
how the cell values with be binned by value. Bin boundaries may optionally be specified
by the following keywords to override the default values for each icon set:

- `min_type`  = "percent" (default), "percentile", "num" or "formula"
- `min_val`     (default: "33" (3 icons), "25" (4 icons) or "20" (5 icons))
- `mid_type`  = "percent" (default), "percentile", "num" or "formula"
- `mid_val`     (default: "50" (4 icons), "40" (5 icons))
- `mid2_type` = "percent" (default), "percentile", "num" or "formula"
- `mid2_val`    (default: "60" (5 icons))
- `max_type`  = "percent" (default), "percentile", "num" or "formula"
- `max_val`     (default: "67" (3 icons), "75" (4 icons) or "80" (5 icons))

The keywords `min_val`, `mid_val`, `mid2_val` and `max_val` may contain numbers (as strings) 
or valid cell references. If `formula` is specified for the related type keyword, a valid 
Excel formula can be provided to evaluate to the bin threshold value to be used.
Three-icon sets require two thresholds (`min_type`/`min_val` and `max_type`/`max_val`), 
four-icon sets require three thresholds (with the addition of `mid_type`/`mid_val`) and 
five-icon sets require four thresholds (adding `mid2_type`/`mid2_val`). Thresholds defined 
(using val and type keywords) that are unnecessary are simply ignored.

Each value can be tested using `>=` (default) or `>`. To change from the default,
optionally set `min_gte`, `mid_gte`, `mid2_gte` and/or `max_gte` to `"false"` to 
use `>` in the comparison. Any other value for these gte keywords will be ignored 
and the default `>=` comparison used.

The built-in icon sets Excel provides are composed of 52 individual icons. It is 
possible to mix and match any of these to make a custom 3-icon, 4-icon or 5-icon 
set by specifying `iconset = "Custom"`. The number of icons in the set will be 
determined by whether the `mid_val`/`mid_type` keywords and `mid2_val`/`mid2_type` 
keywords are provided.

The icons that will be used in a `Custom` iconset are defined using the `icon_list` 
keyword which takes a vector of integers in the range from 1 to 52. For a key relating
integers to the icons they represent, see the [Icon Set](@ref) section in the Formatting 
Guide.

The order in which the symbols are appiled can be reversed from the default order (or, for 
`Custom` icon sets, the order given in `icon_list`), by optionally setting `reverse = "true"`. 
Any other value provided for `reverse` will be ignored, and the default order applied.

The cell value can be suppressed, so that only the icon is shown in the Excel cell by 
optionally specifying `showVal = "false"`. Any other value provided for `showVal` will be 
ignored, and the cell value will be displayed with the icon.

# Examples
```julia
XLSX.setConditionalFormat(s, "F2:F11", :iconSet; iconset="3Arrows")

XLSX.setConditionalFormat(s, 2, :, :iconSet; iconset = "5Boxes",
            reverse = "true",
            showVal = "false",
            min_type="num",  mid_type="percentile", mid2_type="percentile", max_type="num",
            min_val="3",     mid_val="45",          mid2_val="65",          max_val="8",
            min_gte="false", mid_gte="false",       mid2_gte="false",       max_gte="false")

XLSX.setConditionalFormat(s, "A2:A11", :iconSet;
        iconset = "Custom",
        icon_list = [31,24],
        min_type="num",  max_type="formula",
        min_val="3",     max_val="if(\$G\$4=\"y\", \$G\$1+5, 10)")

```

"""
function setConditionalFormat(f, r, type::Symbol; kw...)
    _allkws = Dict{Symbol,Any}(k => v for (k, v) in kw)
    haskey(_allkws, :priority) && throw(XLSXError("Invalid keyword argument: `priority`. Conditional format rules are prioritised in the order they are applied."))
    if type == :colorScale
        setCfColorScale(f, r; allkws=_allkws)
    elseif type == :cellIs
        setCfCellIs(f, r; allkws=_allkws)
    elseif type == :top10
        setCfTop10(f, r; allkws=_allkws)
    elseif type == :aboveAverage
        setCfAboveAverage(f, r; allkws=_allkws)
    elseif type == :timePeriod
        setCfTimePeriod(f, r; allkws=_allkws)
    elseif type ∈ [:containsText, :notContainsText, :beginsWith, :endsWith]
        setCfContainsText(f, r; allkws=_allkws)
    elseif type ∈ [:containsBlanks, :notContainsBlanks, :containsErrors, :notContainsErrors, :duplicateValues, :uniqueValues]
        push!(_allkws, :operator => string(type))
        setCfContainsBlankErrorUniqDup(f, r; allkws=_allkws)
    elseif type == :expression
        setCfFormula(f, r; allkws=_allkws)
    elseif type == :iconSet
        setCfIconSet(f, r; allkws=_allkws)
    elseif type == :dataBar
        setCfDataBar(f, r; allkws=_allkws)
    else
        throw(XLSXError("Invalid conditional format type: $type."))
    end
end

function setConditionalFormat(f, r, c, type::Symbol; kw...)
    _allkws = Dict{Symbol,Any}(k => v for (k, v) in kw)
    haskey(_allkws, :priority) && throw(XLSXError("Invalid keyword argument: `priority`. Conditional format rules are prioritised in the order they are applied."))
    if type == :colorScale
        setCfColorScale(f, r, c; allkws=_allkws)
    elseif type == :cellIs
        setCfCellIs(f, r, c; allkws=_allkws)
    elseif type == :top10
        setCfTop10(f, r, c; allkws=_allkws)
    elseif type == :aboveAverage
        setCfAboveAverage(f, r, c; allkws=_allkws)
    elseif type == :timePeriod
        setCfTimePeriod(f, r, c; allkws=_allkws)
    elseif type ∈ [:containsText, :notContainsText, :beginsWith, :endsWith]
        setCfContainsText(f, r, c; allkws=_allkws)
    elseif type ∈ [:containsBlanks, :notContainsBlanks, :containsErrors, :notContainsErrors, :duplicateValues, :uniqueValues]
        push!(_allkws, :operator => string(type))
        setCfContainsBlankErrorUniqDup(f, r, c; allkws=_allkws)
    elseif type == :expression
        setCfFormula(f, r, c; allkws=_allkws)
    elseif type == :iconSet
        setCfIconSet(f, r, c; allkws=_allkws)
    elseif type == :dataBar
        setCfDataBar(f, r, c; allkws=_allkws)
    else
        throw(XLSXError("Invalid conditional format type: $type."))
    end
end

setCfCellIs(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfCellIs, ws, row, nothing; kw...)
setCfCellIs(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfCellIs, ws, nothing, col; kw...)
setCfCellIs(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfCellIs, ws, nothing, nothing; kw...)
setCfCellIs(ws::Worksheet, ::Colon; kw...) = process_colon(setCfCellIs, ws, nothing, nothing; kw...)
setCfCellIs(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfCellIs(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfCellIs(ws::Worksheet, cell::CellRef; kw...) = setCfCellIs(ws, CellRange(cell, cell); kw...)
setCfCellIs(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfCellIs(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfCellIs(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfCellIs(ws, rng.rng; kw...)
setCfCellIs(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfCellIs(ws, rng.colrng; kw...)
setCfCellIs(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfCellIs(ws, rng.rowrng; kw...)
setCfCellIs(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfCellIs, ws, rng; kw...)
setCfCellIs(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfCellIs, ws, rng; kw...)
setCfCellIs(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfCellIs, xl, sheetcell; kw...)
setCfCellIs(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfCellIs, ws, ref_or_rng; kw...)
setCfCellIs(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfCellIs, ws, row, nothing; kw...)
setCfCellIs(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfCellIs, ws, nothing, col; kw...)
setCfCellIs(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfCellIs, ws, row, col; kw...)
setCfCellIs(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfCellIs, ws, row, col; kw...)
setCfCellIs(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfCellIs, ws, row, col; kw...)
function setCfCellIs(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="greaterThan"
    value::Union{Nothing,String}=nothing
    value2::Union{Nothing,String}=nothing
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing

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

    allcfs    = allCfs(ws)

    !isnothing(value) && !is_valid_cellname(value) && !is_valid_fixed_cellname(value) && isnothing(tryparse(Float64, value)) && throw(XLSXError("Invalid `value`: $value. Must be a number or a CellRef."))
    !isnothing(value2) && !is_valid_cellname(value2) && !is_valid_fixed_cellname(value2) && isnothing(tryparse(Float64, value2)) && throw(XLSXError("Invalid `value2`: $value2. Must be a number or a CellRef."))
    !isnothing(operator) && operator ∉ ["greaterThan", "greaterEqual", "lessThan", "lessEqual", "equal", "notEqual", "between", "notBetween"] && throw(XLSXError("Invalid `operator`: $operator. Valid options are: `greaterThan`, `greaterEqual`, `lessThan`, `lessEqual`, `equal`, `notEqual`, `between`, `notBetween`"))
    wb = get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    if isnothing(value)
        value = all(ismissing.(ws[rng])) ? nothing : string(sum(skipmissing(ws[rng])) / count(!ismissing, ws[rng]))
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

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfContainsText(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfContainsText, ws, row, nothing; kw...)
setCfContainsText(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfContainsText, ws, nothing, col; kw...)
setCfContainsText(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfContainsText, ws, nothing, nothing; kw...)
setCfContainsText(ws::Worksheet, ::Colon; kw...) = process_colon(setCfContainsText, ws, nothing, nothing; kw...)
setCfContainsText(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfContainsText(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfContainsText(ws::Worksheet, cell::CellRef; kw...) = setCfContainsText(ws, CellRange(cell, cell); kw...)
setCfContainsText(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfContainsText(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfContainsText(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsText(ws, rng.rng; kw...)
setCfContainsText(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsText(ws, rng.colrng; kw...)
setCfContainsText(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsText(ws, rng.rowrng; kw...)
setCfContainsText(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfContainsText, ws, rng; kw...)
setCfContainsText(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfContainsText, ws, rng; kw...)
setCfContainsText(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfContainsText, xl, sheetcell; kw...)
setCfContainsText(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfContainsText, ws, ref_or_rng; kw...)
setCfContainsText(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfContainsText, ws, row, nothing; kw...)
setCfContainsText(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfContainsText, ws, nothing, col; kw...)
setCfContainsText(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfContainsText, ws, row, col; kw...)
setCfContainsText(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfContainsText, ws, row, col; kw...)
setCfContainsText(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfContainsText, ws, row, col; kw...)
function setCfContainsText(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="containsText"
    value::Union{Nothing,String}=nothing
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
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

    allcfs    = allCfs(ws)

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

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfTop10(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfTop10, ws, row, nothing; kw...)
setCfTop10(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfTop10, ws, nothing, col; kw...)
setCfTop10(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfTop10, ws, nothing, nothing; kw...)
setCfTop10(ws::Worksheet, ::Colon; kw...) = process_colon(setCfTop10, ws, nothing, nothing; kw...)
setCfTop10(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfTop10(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfTop10(ws::Worksheet, cell::CellRef; kw...) = setCfTop10(ws, CellRange(cell, cell); kw...)
setCfTop10(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfTop10(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfTop10(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfTop10(ws, rng.rng; kw...)
setCfTop10(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfTop10(ws, rng.colrng; kw...)
setCfTop10(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfTop10(ws, rng.rowrng; kw...)
setCfTop10(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfTop10, ws, rng; kw...)
setCfTop10(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfTop10, ws, rng; kw...)
setCfTop10(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfTop10, xl, sheetcell; kw...)
setCfTop10(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfTop10, ws, ref_or_rng; kw...)
setCfTop10(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfTop10, ws, row, nothing; kw...)
setCfTop10(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfTop10, ws, nothing, col; kw...)
setCfTop10(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfTop10, ws, row, col; kw...)
setCfTop10(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfTop10, ws, row, col; kw...)
setCfTop10(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfTop10, ws, row, col; kw...)
function setCfTop10(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="topN"
    value::Union{Nothing,String}="10"
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
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

    allcfs    = allCfs(ws)

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

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfAboveAverage(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfAboveAverage, ws, row, nothing; kw...)
setCfAboveAverage(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfAboveAverage, ws, nothing, col; kw...)
setCfAboveAverage(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfAboveAverage, ws, nothing, nothing; kw...)
setCfAboveAverage(ws::Worksheet, ::Colon; kw...) = process_colon(setCfAboveAverage, ws, nothing, nothing; kw...)
setCfAboveAverage(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfAboveAverage(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfAboveAverage(ws::Worksheet, cell::CellRef; kw...) = setCfAboveAverage(ws, CellRange(cell, cell); kw...)
setCfAboveAverage(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfAboveAverage(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfAboveAverage(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfAboveAverage(ws, rng.rng; kw...)
setCfAboveAverage(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfAboveAverage(ws, rng.colrng; kw...)
setCfAboveAverage(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfAboveAverage(ws, rng.rowrng; kw...)
setCfAboveAverage(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfAboveAverage, ws, rng; kw...)
setCfAboveAverage(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfAboveAverage, ws, rng; kw...)
setCfAboveAverage(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfAboveAverage, xl, sheetcell; kw...)
setCfAboveAverage(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfAboveAverage, ws, ref_or_rng; kw...)
setCfAboveAverage(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfAboveAverage, ws, row, nothing; kw...)
setCfAboveAverage(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfAboveAverage, ws, nothing, col; kw...)
setCfAboveAverage(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfAboveAverage, ws, row, col; kw...)
setCfAboveAverage(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfAboveAverage, ws, row, col; kw...)
setCfAboveAverage(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfAboveAverage, ws, row, col; kw...)
function setCfAboveAverage(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="aboveAverage"
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
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

    allcfs    = allCfs(ws)

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

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfTimePeriod(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfTimePeriod, ws, row, nothing; kw...)
setCfTimePeriod(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfTimePeriod, ws, nothing, col; kw...)
setCfTimePeriod(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfTimePeriod, ws, nothing, nothing; kw...)
setCfTimePeriod(ws::Worksheet, ::Colon; kw...) = process_colon(setCfTimePeriod, ws, nothing, nothing; kw...)
setCfTimePeriod(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfTimePeriod(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfTimePeriod(ws::Worksheet, cell::CellRef; kw...) = setCfTimePeriod(ws, CellRange(cell, cell); kw...)
setCfTimePeriod(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfTimePeriod(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfTimePeriod(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfTimePeriod(ws, rng.rng; kw...)
setCfTimePeriod(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfTimePeriod(ws, rng.colrng; kw...)
setCfTimePeriod(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfTimePeriod(ws, rng.rowrng; kw...)
setCfTimePeriod(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfTimePeriod, ws, rng; kw...)
setCfTimePeriod(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfTimePeriod, ws, rng; kw...)
setCfTimePeriod(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfTimePeriod, xl, sheetcell; kw...)
setCfTimePeriod(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfTimePeriod, ws, ref_or_rng; kw...)
setCfTimePeriod(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfTimePeriod, ws, row, nothing; kw...)
setCfTimePeriod(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfTimePeriod, ws, nothing, col; kw...)
setCfTimePeriod(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfTimePeriod, ws, row, col; kw...)
setCfTimePeriod(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfTimePeriod, ws, row, col; kw...)
setCfTimePeriod(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfTimePeriod, ws, row, col; kw...)
function setCfTimePeriod(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="last7Days"
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
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

    allcfs    = allCfs(ws)

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

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfContainsBlankErrorUniqDup(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfContainsBlankErrorUniqDup, ws, row, nothing; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfContainsBlankErrorUniqDup, ws, nothing, col; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfContainsBlankErrorUniqDup, ws, nothing, nothing; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ::Colon; kw...) = process_colon(setCfContainsBlankErrorUniqDup, ws, nothing, nothing; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfContainsBlankErrorUniqDup(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, cell::CellRef; kw...) = setCfContainsBlankErrorUniqDup(ws, CellRange(cell, cell); kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfContainsBlankErrorUniqDup(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsBlankErrorUniqDup(ws, rng.rng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsBlankErrorUniqDup(ws, rng.colrng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsBlankErrorUniqDup(ws, rng.rowrng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfContainsBlankErrorUniqDup, ws, rng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfContainsBlankErrorUniqDup, ws, rng; kw...)
setCfContainsBlankErrorUniqDup(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfContainsBlankErrorUniqDup, xl, sheetcell; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfContainsBlankErrorUniqDup, ws, ref_or_rng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfContainsBlankErrorUniqDup, ws, row, nothing; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfContainsBlankErrorUniqDup, ws, nothing, col; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfContainsBlankErrorUniqDup, ws, row, col; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfContainsBlankErrorUniqDup, ws, row, col; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfContainsBlankErrorUniqDup, ws, row, col; kw...)
function setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    operator::Union{Nothing,String}="containsBlanks"
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
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

    allcfs    = allCfs(ws)

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

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfFormula, ws, row, nothing; kw...)
setCfFormula(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfFormula, ws, nothing, col; kw...)
setCfFormula(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfFormula, ws, nothing, nothing; kw...)
setCfFormula(ws::Worksheet, ::Colon; kw...) = process_colon(setCfFormula, ws, nothing, nothing; kw...)
setCfFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfFormula(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfFormula(ws::Worksheet, cell::CellRef; kw...) = setCfFormula(ws, CellRange(cell, cell); kw...)
setCfFormula(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfFormula(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfFormula(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfFormula(ws, rng.rng; kw...)
setCfFormula(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfFormula(ws, rng.colrng; kw...)
setCfFormula(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfFormula(ws, rng.rowrng; kw...)
setCfFormula(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfFormula, ws, rng; kw...)
setCfFormula(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfFormula, ws, rng; kw...)
setCfFormula(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfFormula, xl, sheetcell; kw...)
setCfFormula(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfFormula, ws, ref_or_rng; kw...)
setCfFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfFormula, ws, row, nothing; kw...)
setCfFormula(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfFormula, ws, nothing, col; kw...)
setCfFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfFormula, ws, row, col; kw...)
setCfFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfFormula, ws, row, col; kw...)
setCfFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfFormula, ws, row, col; kw...)
function setCfFormula(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    formula::Union{Nothing,String}=nothing
    stopIfTrue::Union{Nothing,String}=nothing
    dxStyle::Union{Nothing,String}=nothing
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
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

    allcfs    = allCfs(ws)

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

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfColorScale(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfColorScale, ws, row, nothing; kw...)
setCfColorScale(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfColorScale, ws, nothing, col; kw...)
setCfColorScale(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfColorScale, ws, nothing, nothing; kw...)
setCfColorScale(ws::Worksheet, ::Colon; kw...) = process_colon(setCfColorScale, ws, nothing, nothing; kw...)
setCfColorScale(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfColorScale(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfColorScale(ws::Worksheet, cell::CellRef; kw...) = setCfColorScale(ws, CellRange(cell, cell); kw...)
setCfColorScale(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfColorScale(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfColorScale(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.rng; kw...)
setCfColorScale(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.colrng; kw...)
setCfColorScale(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.rowrng; kw...)
setCfColorScale(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfColorScale, ws, rng; kw...)
setCfColorScale(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfColorScale, ws, rng; kw...)
setCfColorScale(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfColorScale, xl, sheetcell; kw...)
setCfColorScale(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfColorScale, ws, ref_or_rng; kw...)
setCfColorScale(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfColorScale, ws, row, nothing; kw...)
setCfColorScale(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfColorScale, ws, nothing, col; kw...)
setCfColorScale(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfColorScale, ws, row, col; kw...)
setCfColorScale(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfColorScale, ws, row, col; kw...)
setCfColorScale(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfColorScale, ws, row, col; kw...)
function setCfColorScale(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int
    colorscale::Union{Nothing,String}=nothing
    min_type::Union{Nothing,String}="min"
    min_val::Union{Nothing,String}=nothing
    min_col::Union{Nothing,String}="FFF8696B"
    mid_type::Union{Nothing,String}=nothing
    mid_val::Union{Nothing,String}=nothing
    mid_col::Union{Nothing,String}="FFFCFCFF"
    max_type::Union{Nothing,String}="max"
    max_val::Union{Nothing,String}=nothing
    max_col::Union{Nothing,String}="FFFFEB84"

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

    allcfs    = allCfs(ws)

    let new_pr, new_cf

        new_pr = string(next_cf_priority!(ws))

        if isnothing(colorscale)

            min_type in ["min", "percentile", "percent", "num", "formula"] || throw(XLSXError("Invalid min_type: $min_type. Valid options are: min, percentile, percent, num, formula."))
            if min_type == "min"
                min_val = nothing
            end
            min_type == "formula" || isnothing(min_val) || is_valid_fixed_cellname(min_val) || is_valid_fixed_sheet_cellname(min_val) || !isnothing(tryparse(Float64, min_val)) || throw(XLSXError("Invalid min_val: `$min_val`. Valid options (unless min_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))
 
            isnothing(mid_type) || mid_type in ["percentile", "percent", "num", "formula"] || throw(XLSXError("Invalid mid_type: $mid_type. Valid options are: percentile, percent, num, formula."))
            (!isnothing(mid_type) && mid_type == "formula") || isnothing(mid_val) || is_valid_fixed_cellname(mid_val) || is_valid_fixed_sheet_cellname(mid_val) || !isnothing(tryparse(Float64, mid_val)) || throw(XLSXError("Invalid mid_val: `$mid_val`. Valid options (unless mid_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))

            max_type in ["max", "percentile", "percent", "num", "formula"] || throw(XLSXError("Invalid max_type: $max_type. Valid options are: max, percentile, percent, num, formula."))
            if max_type == "max"
                max_val = nothing
            end
            max_type == "formula" || isnothing(max_val) || is_valid_fixed_cellname(max_val) || is_valid_fixed_sheet_cellname(max_val) || !isnothing(tryparse(Float64, max_val)) || throw(XLSXError("Invalid max_val: `$max_val`. Valid options (unless max_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))

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

            push!(csc, XML.Element("$(pfx)color"; rgb=get_color(min_col)))
            !isnothing(mid_type) && push!(csc, XML.Element("$(pfx)color"; rgb=get_color(mid_col)))
            push!(csc, XML.Element("$(pfx)color"; rgb=get_color(max_col)))                
            push!(cfx, csc)
        else
            if !haskey(colorscales, colorscale)
                throw(XLSXError("Invalid colorscale option chosen: $colorscale. Valid options are: $(keys(colorscales))."))
            end
            cfx = copynode(colorscales[colorscale])
            cfx["priority"] = new_pr
        end

        update_worksheet_cfx!(allcfs, cfx, ws, rng)

    end

    return 0
end

setCfIconSet(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfIconSet, ws, row, nothing; kw...)
setCfIconSet(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfIconSet, ws, nothing, col; kw...)
setCfIconSet(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfIconSet, ws, nothing, nothing; kw...)
setCfIconSet(ws::Worksheet, ::Colon; kw...) = process_colon(setCfIconSet, ws, nothing, nothing; kw...)
setCfIconSet(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfIconSet(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfIconSet(ws::Worksheet, cell::CellRef; kw...) = setCfIconSet(ws, CellRange(cell, cell); kw...)
setCfIconSet(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfIconSet(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfIconSet(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfIconSet(ws, rng.rng; kw...)
setCfIconSet(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfIconSet(ws, rng.colrng; kw...)
setCfIconSet(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfIconSet(ws, rng.rowrng; kw...)
setCfIconSet(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfIconSet, ws, rng; kw...)
setCfIconSet(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfIconSet, ws, rng; kw...)
setCfIconSet(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfIconSet, xl, sheetcell; kw...)
setCfIconSet(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfIconSet, ws, ref_or_rng; kw...)
setCfIconSet(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfIconSet, ws, row, nothing; kw...)
setCfIconSet(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfIconSet, ws, nothing, col; kw...)
setCfIconSet(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfIconSet, ws, row, col; kw...)
setCfIconSet(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfIconSet, ws, row, col; kw...)
setCfIconSet(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfIconSet, ws, row, col; kw...)
function setCfIconSet(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    iconset::Union{Nothing,String} = "3TrafficLights"
    reverse::Union{Nothing,String} = nothing
    showVal::Union{Nothing,String} = nothing
    min_type::Union{Nothing,String} = nothing
    min_val::Union{Nothing,String} = nothing
    min_gte::Union{Nothing,String} = nothing
    mid_type::Union{Nothing,String} = nothing
    mid_val::Union{Nothing,String} = nothing
    mid_gte::Union{Nothing,String} = nothing
    mid2_type::Union{Nothing,String} = nothing
    mid2_val::Union{Nothing,String} = nothing
    mid2_gte::Union{Nothing,String} = nothing
    max_type::Union{Nothing,String} = nothing
    max_val::Union{Nothing,String} = nothing
    max_gte::Union{Nothing,String} = nothing
    icon_list::Union{Nothing,Vector{Int64}}=nothing

    for (k, v) in allkws
        if k == :iconset
            iconset = v
        elseif k == :reverse
            reverse = v
        elseif k == :showVal
            showVal = v
        elseif k == :min_type
            min_type = v
        elseif k == :min_val
            min_val = v
        elseif k == :min_gte
            min_gte = v 
        elseif k == :mid_type
            mid_type = v
        elseif k == :mid_val
            mid_val = v
        elseif k == :mid_gte
            mid_gte = v
        elseif k == :mid2_type
            mid2_type = v
        elseif k == :mid2_val
            mid2_val = v
        elseif k == :mid2_gte
            mid2_gte = v
        elseif k == :max_type
            max_type = v
        elseif k == :max_val
            max_val = v
        elseif k == :max_gte
            max_gte = v
        elseif k == :icon_list
            icon_list = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `iconset`, `reverse`, `showVal`, `min_type`, `min_val`, `min_gte`, `mid_type`, `mid_val`, `mid_gte`, `mid2_type`, `mid2_val`, `mid2_gte`, `max_type`, `max_val`, `max_gte`, `icon_list`."))
        end
    end
    
    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"

    allcfs    = allCfs(ws)
    allextcfs = allExtCfs(ws)

    let new_pr, new_cf

        new_pr = string(next_cf_priority!(ws))

        isnothing(min_type) || min_type in ["percentile", "percent", "num", "formula"] || throw(XLSXError("Invalid min_type: $min_type. Valid options are: percentile, percent, num, formula."))
        (!isnothing(min_type) && min_type == "formula") || isnothing(min_val) || is_valid_fixed_cellname(min_val) || is_valid_fixed_sheet_cellname(min_val) || !isnothing(tryparse(Float64, min_val)) || throw(XLSXError("Invalid min_val: `$min_val`. Valid options (unless min_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))
        isnothing(mid_type) || mid_type in ["percentile", "percent", "num", "formula"] || throw(XLSXError("Invalid mid_type: $mid_type. Valid options are: percentile, percent, num, formula."))
        (!isnothing(mid_type) && mid_type == "formula") || isnothing(mid_val) || is_valid_fixed_cellname(mid_val) || !is_valid_fixed_sheet_cellname(mid_val) || !isnothing(tryparse(Float64, mid_val)) || throw(XLSXError("Invalid mid_val: `$mid_val`. Valid options (unless mid_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))
        isnothing(mid2_type) || mid2_type in ["percentile", "percent", "num", "formula"] || throw(XLSXError("Invalid mid_type: $mid2_type. Valid options are: percentile, percent, num, formula."))
        (!isnothing(mid2_type) && mid2_type == "formula") || isnothing(mid2_val) || is_valid_fixed_cellname(mid2_val) || is_valid_fixed_sheet_cellname(mid2_val) || !isnothing(tryparse(Float64, mid2_val)) || throw(XLSXError("Invalid mid2_type: `$mid2_val`. Valid options (unless mid2_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))
        isnothing(max_type) || max_type in ["percentile", "percent", "num", "formula"] || throw(XLSXError("Invalid max_type: $max_type. Valid options are: percentile, percent, num, formula."))
        (!isnothing(max_type) && max_type == "formula") || isnothing(max_val) || is_valid_fixed_cellname(max_val) || is_valid_fixed_sheet_cellname(max_val) || !isnothing(tryparse(Float64, max_val)) || throw(XLSXError("Invalid max_val: `$max_val`. Valid options (unless max_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))

        min_val  = _normalise_cfvo_val(ws, min_val)
        mid_val  = _normalise_cfvo_val(ws, mid_val)
        mid2_val = _normalise_cfvo_val(ws, mid2_val)
        max_val  = _normalise_cfvo_val(ws, max_val)
        
        if !haskey(iconsets, iconset)
            throw(XLSXError("Invalid iconset option chosen: $iconset. Valid options are: $(keys(iconsets))"))
        end
        l = first(iconset)
        cfx = copynode(iconsets[iconset])
        if l == 'C'
            cfvo = XML.Element("x14:cfvo", type="percent")
            push!(cfvo, XML.Element("xm:f", XML.Text("dummy")))
            push!(cfx[1], copynode(cfvo)) # for min_val
            push!(cfx[1], copynode(cfvo)) # for max_val
            if isnothing(min_type) || isnothing(min_val) || isnothing(max_type) || isnothing(max_val)
                throw(XLSXError("No type or val keywords defined. Must define at least `min_type`, `min_val`, `max_type` and `max_val` for a custom iconSet"))
            elseif isnothing(mid_type) || isnothing(mid_val)
                list = [(min_type, min_val, min_gte), (max_type, max_val, max_gte)]
                nicons = 3
            elseif isnothing(mid2_type) || isnothing(mid2_val)
                push!(cfx[1], copynode(cfvo)) # for mid_val
                cfx[1]["iconSet"] = "4Arrows"
                nicons = 4
                list = [(min_type, min_val, min_gte), (mid_type, mid_val, mid_gte), (max_type, max_val, max_gte)]
            else
                push!(cfx[1], copynode(cfvo)) # for mid_val
                push!(cfx[1], copynode(cfvo)) # for mid2_val
                cfx[1]["iconSet"] = "5Quarters"
                nicons = 5
                list = [(min_type, min_val, min_gte), (mid_type, mid_val, mid_gte), (mid2_type, mid2_val, mid2_gte), (max_type, max_val, max_gte)]
            end
        elseif l == '5'
            list = [(min_type, min_val, min_gte), (mid_type, mid_val, mid_gte), (mid2_type, mid2_val, mid2_gte), (max_type, max_val, max_gte)]
        elseif l == '4'
            list = [(min_type, min_val, min_gte), (mid_type, mid_val, mid_gte), (max_type, max_val, max_gte)]
        else
            list = [(min_type, min_val, min_gte), (max_type, max_val, max_gte)]
        end
        if iconset in ["3Triangles", "3Stars", "5Boxes", "Custom"]
            cfx["id"] = "{" * uppercase(string(UUIDs.uuid4(ws.package.uuid_rng))) * "}"
            cfx["priority"] = new_pr
            if !isnothing(showVal) && showVal == "false"
                cfx[1]["showValue"] = "0"
            end
            if !isnothing(reverse) && reverse == "true"
                if iconset == "Custom"
                    reverse!(icon_list)
                else
                    cfx[1]["reverse"] = "1"
                end
            end
            for (i, (type, val, gte)) in enumerate(list)
                if !isnothing(type)
                    cfx[1][i+1]["type"] = type # Need +1 because the first <cfvo> is always 0 percent.
                end
                if !isnothing(val)
                    if !isnothing(type) && type == "formula"
                        c = XML.Element("xm:f", XML.Text("(" * val * ")"))
                    else
                        c = XML.Element("xm:f", XML.Text(val))
                    end
                    if isempty(XML.children(cfx[1][i+1]))
                        old = cfx[1][i+1]
                        cfx[1][i+1] = XML.Node{String}(XML.Element, XML.tag(old), old.attributes, nothing, XML.Node{String}[c])
                    else
                        cfx[1][i+1][1] = c
                    end
                end
                if !isnothing(gte) && gte == "false"
                    cfx[1][i+1]["gte"] = "0"
                end
            end
            if iconset == "Custom"
                if isnothing(icon_list)
                    throw(XLSXError("No custom icons specified. Must specify between two and four icons."))
                elseif length(icon_list) < nicons
                    throw(XLSXError("Too few custom icons specified: $(length(icon_list)). Expected $nicons"))
                end
                for (count, icon) in enumerate(string.(icon_list))
                    if !isnothing(icon)
                        if !haskey(allIcons, icon)
                            throw(XLSXError("Invalid custom icon specified: $icon. Valid values are \"1\" to \"52\"."))
                        end
                        i = allIcons[icon]
                        push!(cfx[1], XML.Element("x14:cfIcon", iconSet=first(i), iconId=last(i)))
                        count == nicons && break
                    end
                end
            end
            update_worksheet_ext_cfx!(allextcfs, cfx, ws, rng)
        else
            cfx["priority"] = new_pr
            if !isnothing(showVal) && showVal == "false"
                cfx[1]["showValue"] = "0"
            end
            if !isnothing(reverse) && reverse == "true"
                cfx[1]["reverse"] = "1"
            end
            for (i, (type, val, gte)) in enumerate(list)
                if !isnothing(val)
                    if !isnothing(type) && type == "formula"
                        cfx[1][i+1]["val"] = "(" * val * ")"
                    else
                        cfx[1][i+1]["val"] = val
                    end
                end
                if !isnothing(type)
                    cfx[1][i+1]["type"] = type
                end
                if !isnothing(gte) && gte == "false"
                    cfx[1][i+1]["gte"] = "0"
                end
            end
            update_worksheet_cfx!(allcfs, cfx, ws, rng)
        end

    end

    return 0
end

setCfDataBar(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfDataBar, ws, row, nothing; kw...)
setCfDataBar(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfDataBar, ws, nothing, col; kw...)
setCfDataBar(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfDataBar, ws, nothing, nothing; kw...)
setCfDataBar(ws::Worksheet, ::Colon; kw...) = process_colon(setCfDataBar, ws, nothing, nothing; kw...)
setCfDataBar(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfDataBar(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfDataBar(ws::Worksheet, cell::CellRef; kw...) = setCfDataBar(ws, CellRange(cell, cell); kw...)
setCfDataBar(ws::Worksheet, cell::SheetCellRef; kw...) = do_sheet_names_match(ws, cell) && setCfDataBar(ws, CellRange(cell.cellref, cell.cellref); kw...)
setCfDataBar(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfDataBar(ws, rng.rng; kw...)
setCfDataBar(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfDataBar(ws, rng.colrng; kw...)
setCfDataBar(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfDataBar(ws, rng.rowrng; kw...)
setCfDataBar(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfDataBar, ws, rng; kw...)
setCfDataBar(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfDataBar, ws, rng; kw...)
setCfDataBar(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfDataBar, xl, sheetcell; kw...)
setCfDataBar(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfDataBar, ws, ref_or_rng; kw...)
setCfDataBar(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_cf_veccolon(setCfDataBar, ws, row, nothing; kw...)
setCfDataBar(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_veccolon(setCfDataBar, ws, nothing, col; kw...)
setCfDataBar(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfDataBar, ws, row, col; kw...)
setCfDataBar(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_cf_vecint(setCfDataBar, ws, row, col; kw...)
setCfDataBar(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_cf_vecint(setCfDataBar, ws, row, col; kw...)
function setCfDataBar(ws::Worksheet, rng::CfRange; allkws::Dict{Symbol,Any}=())::Int

    databar::Union{Nothing,String}="bluegrad"
    priority::Union{Nothing,Int}=nothing  # internal only - see `new_pr` below
    showVal::Union{Nothing,String}=nothing
    gradient::Union{Nothing,String}=nothing
    borders::Union{Nothing,String}=nothing
    sameNegFill::Union{Nothing,String}=nothing
    sameNegBorders::Union{Nothing,String}=nothing
    direction::Union{Nothing,String}=nothing
    axis_pos::Union{Nothing,String}=nothing
    axis_col::Union{Nothing,String}=nothing
    min_type::Union{Nothing,String}=nothing
    min_val::Union{Nothing,String}=nothing
    max_type::Union{Nothing,String}=nothing
    max_val::Union{Nothing,String}=nothing
    fill_col::Union{Nothing,String}=nothing
    border_col::Union{Nothing,String}=nothing
    neg_fill_col::Union{Nothing,String}=nothing
    neg_border_col::Union{Nothing,String}=nothing

    for (k, v) in allkws
        if k == :databar
            databar = v
        elseif k == :priority
            priority = v
        elseif k == :showVal
            showVal = v
        elseif k == :gradient
            gradient = v
        elseif k == :borders
            borders = v
        elseif k == :sameNegFill
            sameNegFill = v
        elseif k == :sameNegBorders
            sameNegBorders = v
        elseif k == :direction
            direction = v
        elseif k == :axis_pos
            axis_pos = v
        elseif k == :axis_col
            axis_col = v
        elseif k == :min_type
            min_type = v
        elseif k == :min_val
            min_val = v
        elseif k == :max_type
            max_type = v
        elseif k == :max_val
            max_val = v
        elseif k == :fill_col
            fill_col = v
        elseif k == :border_col
            border_col = v
        elseif k == :neg_fill_col
            neg_fill_col = v
        elseif k == :neg_border_col
            neg_border_col = v
        else
            throw(XLSXError("Invalid keyword argument: $k. Valid options are: `databar`, `showVal`, `gradient`, `borders`, `sameNegFill`, `sameNegBorders`, `direction`, `axis_pos`, `axis_col`, `min_type`, `min_val`, `max_type`, `max_val`, `fill_col`, `border_col`, `neg_fill_col`, `neg_border_col`."))
        end
    end

    _check_cf_range(ws, rng)

    pfx = get_prefix(ws)
    pfx = pfx == "" ? pfx : pfx * ":"

    allcfs    = allCfs(ws)
    allextcfs = allExtCfs(ws)

    let new_pr, new_cf

        # `priority` is for internal callers only and is deliberately undocumented: it is
        # omitted from the valid-keyword list above so a user passing it gets an error.
        # `setColoredDataBars` reserves a contiguous block of priorities up front and passes
        # them in band order, so the bands are numbered predictably rather than in whatever
        # order the partition happens to yield. Note it bypasses `next_cf_priority!`, so a
        # caller supplying it owns the number: no uniqueness check, and the counter does not
        # advance.
        new_pr = string(something(priority, next_cf_priority!(ws)))

        isnothing(min_type) || min_type in ["least", "percentile", "percent", "num", "formula", "automatic"] || throw(XLSXError("Invalid min_type: $min_type. Valid options are: least, percentile, percent, num, formula."))
        if min_type in ["least", "automatic"]
            min_val = nothing
        end
        (!isnothing(min_type) && min_type == "formula") || isnothing(min_val) || is_valid_fixed_cellname(min_val) || is_valid_fixed_sheet_cellname(min_val) || !isnothing(tryparse(Float64, min_val)) || throw(XLSXError("Invalid min_val: `$min_val`. Valid options (unless min_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))
        isnothing(max_type) || max_type in ["highest", "percentile", "percent", "num", "formula", "automatic"] || throw(XLSXError("Invalid max_type: $max_type. Valid options are: highest, percentile, percent, num, formula."))
        if max_type in ["highest", "automatic"]
            max_val = nothing
        end
        (!isnothing(max_type) && max_type == "formula") || isnothing(max_val) || is_valid_fixed_cellname(max_val) || is_valid_fixed_sheet_cellname(max_val) || !isnothing(tryparse(Float64, max_val)) || throw(XLSXError("Invalid max_val: `$max_val`. Valid options (unless max_type is `formula`) are a CellRef (e.g. `\$A\$1`) or a number."))

        min_val = _normalise_cfvo_val(ws, min_val)
        max_val = _normalise_cfvo_val(ws, max_val)
        
        if !haskey(databars, databar)
            throw(XLSXError("Invalid dataBar option chosen: $databar. Valid options are: $(keys(databars))"))
        end

        allkws::Dict{String,Union{String,Nothing}} = Dict(
            "showVal" => showVal,
            "gradient" => gradient,
            "borders" => borders,
            "sameNegFill" => sameNegFill,
            "sameNegBorders" => sameNegBorders,
            "direction" => direction,
            "min_type" => min_type,
            "min_val" => min_val,
            "max_type" => max_type,
            "max_val" => max_val,
            "fill_col" => fill_col,
            "border_col" => border_col,
            "neg_fill_col" => neg_fill_col,
            "neg_border_col" => neg_border_col,
            "axis_pos" => axis_pos,
            "axis_col" => axis_col
        )

        for (k, w) in allkws # Allow user input to override any default value
            if isnothing(w)
                if haskey(databars[databar], k)
                    allkws[k] = databars[databar][k]
                end
            end
        end
        for kw in ["showVal", "gradient", "borders", "sameNegFill", "sameNegBorders"]
            haskey(allkws, kw) && isValidKw(kw, allkws[kw], ["true", "false"])
        end
        haskey(allkws, "direction") && isValidKw("direction", allkws["direction"], ["leftToRight", "rightToLeft"])
        haskey(allkws, "axis_pos") && isValidKw("axis_pos", allkws["axis_pos"], ["middle", "none"])

        # Define basic elements of dataBar definition
        id = "{" * uppercase(string(UUIDs.uuid4(ws.package.uuid_rng))) * "}"
        mnt = allkws["min_type"] ∈ ["automatic", "least"] ? "min" : allkws["min_type"]
        mxt = allkws["max_type"] ∈ ["automatic", "highest"] ? "max" : allkws["max_type"]

        cfx = XML.Element("$(pfx)cfRule"; type="dataBar", priority=new_pr)
        cdb = XML.Element("$(pfx)dataBar")
        push!(cdb, isnothing(allkws["min_val"]) ?
                XML.Element("$(pfx)cfvo"; type=mnt) :
                XML.Element("$(pfx)cfvo"; type=mnt, val=allkws["min_val"])
        )

        push!(cdb, isnothing(allkws["max_val"]) ?
                XML.Element("$(pfx)cfvo"; type=mxt) :
                XML.Element("$(pfx)cfvo"; type=mxt, val=allkws["max_val"])
        )

        push!(cdb, XML.Element("$(pfx)color"; rgb=get_color(allkws["fill_col"])))

        push!(cfx, cdb)

        if haskey(allkws, "showVal") && !isnothing(allkws["showVal"]) && allkws["showVal"] == "false"
            cdb["showValue"] = "0"
        end
        # Links the 2007 rule to its x14 counterpart by id. Excel expects the <ext> in an
        # <extLst> inside <cfRule>, as a sibling following <dataBar> - not inside <dataBar>
        # itself, where it is silently discarded on open along with the x14 rule it links to.
        cfx_ext = XML.Element("$(pfx)ext")
        cfx_ext["uri"] = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
        cfx_ext["xmlns:x14"] = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        push!(cfx_ext, XML.Element("x14:id", XML.Text(id)))
        cfx_extlst = XML.Element("$(pfx)extLst")
        push!(cfx_extlst, cfx_ext)
        push!(cfx, cfx_extlst)

        # Define extension elements of dataBar definition
        emnt = allkws["min_type"] == "automatic" ? "autoMin" : allkws["min_type"] == "least" ? "min" : allkws["min_type"]
        emxt = allkws["max_type"] == "automatic" ? "autoMax" : allkws["max_type"] == "highest" ? "max" : allkws["max_type"]
        emnv = allkws["min_type"] == "formula" ? "(" * allkws["min_val"] * ")" : allkws["min_val"]
        emxv = allkws["max_type"] == "formula" ? "(" * allkws["max_val"] * ")" : allkws["max_val"]
        ext_cfx = XML.Element("x14:cfRule", type="dataBar", id=id)
        ext_db = XML.Element("x14:dataBar", minLength="0", maxLength="100")
        valmin = XML.Element("x14:cfvo", type=emnt)
        !isnothing(emnv) && push!(valmin, XML.Element("xm:f", XML.Text(emnv)))
        valmax = XML.Element("x14:cfvo", type=emxt)
        !isnothing(emxv) && push!(valmax, XML.Element("xm:f", XML.Text(emxv)))
        push!(ext_db, valmin)
        push!(ext_db, valmax)
        if allkws["gradient"] == "false"
            ext_db["gradient"] = "0"
        end
        do_borders = haskey(allkws, "borders") && allkws["borders"] == "true"
        if do_borders
            ext_db["border"] = "1"
            if haskey(allkws, "border_col") && !isnothing(allkws["border_col"])
                push!(ext_db, XML.Element("x14:borderColor", rgb=get_color(allkws["border_col"])))
            else
                push!(ext_db, XML.Element("x14:borderColor", rgb=get_color("FF638EC6"))) # Default colour
            end
        end
        if haskey(allkws, "direction")
            if allkws["direction"] == "leftToRight"
                ext_db["direction"] = "leftToRight"
            elseif allkws["direction"] == "rightToLeft"
                ext_db["direction"] = "rightToLeft"
            end
        end
        if haskey(allkws, "sameNegFill") && allkws["sameNegFill"] == "true"
            ext_db["negativeBarColorSameAsPositive"] = "1"
        else
            if haskey(allkws, "neg_fill_col") && !isnothing(allkws["neg_fill_col"])
                push!(ext_db, XML.Element("x14:negativeFillColor", rgb=get_color(allkws["neg_fill_col"])))
            else
                push!(ext_db, XML.Element("x14:negativeFillColor", rgb=get_color("FFFF0000"))) # Default colour
            end
        end
        if do_borders && haskey(allkws, "sameNegBorders") && allkws["sameNegBorders"] == "false"
            ext_db["negativeBarBorderColorSameAsPositive"] = "0"
            if haskey(allkws, "neg_border_col") && !isnothing(allkws["neg_border_col"])
                push!(ext_db, XML.Element("x14:negativeBorderColor", rgb=get_color(allkws["neg_border_col"])))
            else
                push!(ext_db, XML.Element("x14:negativeBorderColor", rgb=get_color("FFFF0000"))) # Default colour
            end
        end
        if haskey(allkws, "axis_pos")
            if allkws["axis_pos"] == "none"
                ext_db["axisPosition"] = "none"
            elseif allkws["axis_pos"] == "middle"
                ext_db["axisPosition"] = "middle"
            end
        end
        haskey(allkws, "axis_col") && push!(ext_db, XML.Element("x14:axisColor", rgb=get_color(allkws["axis_col"])))
        push!(ext_cfx, ext_db)

        update_worksheet_cfx!(allcfs, cfx, ws, rng)            # Add basic elements to worksheet xml file
        update_worksheet_ext_cfx!(allextcfs, ext_cfx, ws, rng) # Add extension elements to worksheet xml file
    end
    return 0
end

# ---------------------------------------------------------------------------
# Non-contiguous range support for conditional formats
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# setColoredDataBars
# ---------------------------------------------------------------------------

"""
    setColoredDataBars(sheet, rng; [bands], [colors], [breaks], [gte], [showVal],
                       [direction], [min_type], [min_val], [max_type], [max_val],
                       [clear]) -> Vector{Pair}

Apply data bars to `rng` whose colour changes with value as well as their length.

Excel has no such conditional format: a `dataBar` rule carries a single colour and
applies to every cell in its range. `setColoredDataBars` splits `rng` into value-based
bands and writes one data bar rule per band, each over the non-contiguous range of
cells falling in that band. The result is a valid workbook that Excel reads without
complaint, but it is not something Excel's own interface can produce, and it is not
maintained by Excel: see the note on editing below.

`bands` (default 5) sets how many bands to use, with boundaries placed at equal
intervals between the smallest and largest value in `rng`. Give `breaks` instead to
place the boundaries yourself, as a sorted vector of `length(bands)-1` values; `breaks`
overrides `bands`. By default a cell whose value equals a break falls in the band above
it; set `gte=false` to place it in the band below, or give a vector of `Bool` with one
entry per break.

`colors` accepts either one colour per band, or exactly two colours to interpolate
between (in LCHab, so intermediate hues stay clean). Colours may be names or 8-digit
`AARRGGBB` strings, as elsewhere in XLSX.jl. Note that Colors.jl's named colormaps are
not accepted; for a red-amber-green ramp give the three colours explicitly.

`showVal` and `direction` are passed through to each band's rule unchanged. Bars are
drawn with a solid fill rather than a gradient.

`min_type`, `min_val`, `max_type` and `max_val` set the value axis shared by every band,
and default to `num` with the minimum and maximum of `rng`. Only `num` and `formula` are
accepted: the relative types (`percent`, `percentile`, `min`, `max`, `automatic`) are
evaluated over each rule's own cells, which for a band is only that band's values, so
every band would span the full bar width and bar lengths would carry no meaning.

By default (clear=true) any conditional format lying entirely within rng is removed
before the new bands are written, so repeated calls replace rather than accumulate. This
matters when re-banding after the data have changed: bands from an earlier call name the
cells they covered then, and left in place they would overlay the new ones. Set
clear=false to add the bands alongside whatever is already there.

Returns the band label => range pairs written. Bands containing no cells are dropped, so
the result may be shorter than bands. Labels keep their position on the scale rather
than being renumbered: banding data with only two distinct values across bands=5 may
return labels 1 and 5, using the first and last colours of the ramp.

!!! note
    `setColoredDataBars` is a static formatting function. The bands reflect cell values 
    at the time of writing. If a value is later edited in Excel its bar length updates, 
    but its color does not, because the cell stays in the band it was originally 
    assigned to. Rewrite the conditional formatting to re-band it.

    This is a limitation of Excel itself, which does not support dataBars whose color 
    varies with value. `setColoredDataBars` is a workaround for this.


!!! note "Experimental"

    This function is experimental. It has a full set of tests and Documentation is 
    published but it is not yet marked as public. It may be better included in 
    future only as a documented example rather than as a package function.

# Examples

```julia
julia> XLSX.setColoredDataBars(sheet, "B2:B101")

julia> XLSX.setColoredDataBars(sheet, "B2:B101"; bands=3, colors=["green", "orange", "red"])

julia> XLSX.setColoredDataBars(sheet, "B2:B101"; breaks=[10, 50], colors=["green", "red"])

julia> XLSX.setColoredDataBars(sheet, "B2:B101"; bands=8, min_val="0", max_val="100")
```

See also [`setConditionalFormat`](@ref), [`clearConditionalFormats`](@ref).
"""
function setColoredDataBars end

setColoredDataBars(ws::Worksheet, rng::AbstractString; kw...) =
    setColoredDataBars(ws, CellRange(rng); kw...)
setColoredDataBars(ws::Worksheet, cell::CellRef; kw...) =
    setColoredDataBars(ws, CellRange(cell, cell); kw...)
setColoredDataBars(ws::Worksheet, rng::SheetCellRange; kw...) =
    do_sheet_names_match(ws, rng) && setColoredDataBars(ws, rng.rng; kw...)

function setColoredDataBars(ws::Worksheet, rng::CellRange;
                            bands::Int=5,
                            colors=["green", "red"],
                            breaks=nothing,
                            gte::Union{Bool,AbstractVector{Bool}}=true,
                            showVal::Union{Nothing,String}=nothing,
                            direction::Union{Nothing,String}=nothing,
                            min_type::Union{Nothing,String}=nothing,
                            min_val::Union{Nothing,String}=nothing,
                            max_type::Union{Nothing,String}=nothing,
                            max_val::Union{Nothing,String}=nothing,
                            clear::Bool=true)

    _check_cf_areas(ws, rng)
    
    clear && clearConditionalFormats(ws, rng)

    vals = filter(v -> v isa Real && !isnan(v), vec(getdata(ws, rng)))
    isempty(vals) && throw(XLSXError("No numeric values in `$rng` to band."))
    lo, hi = float(minimum(vals)), float(maximum(vals))

    brk = if isnothing(breaks)
        bands < 1 && throw(XLSXError("`bands` must be at least 1, got $bands."))
        bands == 1 && return _single_bar(ws, rng, colors, lo, hi,
                                         showVal, direction, min_type, min_val, max_type, max_val)
        hi > lo || throw(XLSXError("All numeric values in `$rng` are equal; cannot form bands."))
        collect(range(lo, hi; length=bands + 1))[2:end-1]
    else
        bands = length(breaks) + 1
        collect(breaks)
    end

    cols = _band_colors(colors, bands)

    for (t, name) in ((min_type, "min_type"), (max_type, "max_type"))
        isnothing(t) || t in ["num", "formula"] || throw(XLSXError(
            "`$name` must be `num` or `formula` for coloured data bars; `$t` is evaluated " *
            "separately for each band, which makes bar lengths incomparable between bands."))
    end
    mnt = something(min_type, "num")
    mxt = something(max_type, "num")
    mnv = something(min_val, string(lo))
    mxv = something(max_val, string(hi))

    p = partition(ws, rng, brk; labels=1:bands, gte=gte)

    base = next_cf_priority!(ws)
    for _ in 2:length(p)
        next_cf_priority!(ws)          # reserve the rest
    end

    for (n, (i, ncr)) in enumerate(p)
        kws = Dict{Symbol,Any}(
            :priority => base + n - 1,
            :fill_col => cols[i],
            :gradient => "false",
            :borders => "false",
            :sameNegFill => "true",
            :min_type => mnt, :min_val => mnv,
            :max_type => mxt, :max_val => mxv,
        )
        isnothing(showVal) || (kws[:showVal] = showVal)
        isnothing(direction) || (kws[:direction] = direction)
        setCfDataBar(ws, ncr; allkws=kws)
    end
    return p
end

# `bands=1` degenerates to an ordinary data bar over the whole contiguous range:
# no partition, no non-contiguous sqref.
function _single_bar(ws, rng, colors, lo, hi, showVal, direction,
                     min_type, min_val, max_type, max_val)
    kws = Dict{Symbol,Any}(
        :fill_col => only(_band_colors(colors, 1)),
        :gradient => "false",
        :borders => "false",
        :sameNegFill => "true",
        :min_type => something(min_type, "num"), :min_val => something(min_val, string(lo)),
        :max_type => something(max_type, "num"), :max_val => something(max_val, string(hi)),
    )
    isnothing(showVal) || (kws[:showVal] = showVal)
    isnothing(direction) || (kws[:direction] = direction)
    setCfDataBar(ws, rng; allkws=kws)
    return [1 => rng]
end

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
    ws_el = sheetdoc[end]
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
    i, j = get_idces(sheetdoc, "worksheet", "extLst")
    if !isnothing(j)
        m, k = get_idces(sheetdoc[i], "extLst", "ext")
        if !isnothing(k)
            o, p = get_idces(sheetdoc[i][j], "ext", "x14:conditionalFormattings")
            if !isnothing(p)
                blk = sheetdoc[i][j][k][p]
                for q in reverse(eachindex(XML.children(blk)))
                    b = blk[q]
                    els = xml_elements(b)
                    s = findlast(e -> localname(e) == "sqref", els)
                    isnothing(s) && continue
                    if _cf_within(ws, XML.simple_value(els[s]), rng)
                        deleteat!(blk.children, q)
                    end
                end
            end
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