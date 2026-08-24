# Check if a string is a valid named color in Colors.jl and convert to "FFRRGGBB" if it is.
get_colorant(color_symb::Symbol) = get_colorant(String(color_symb))
function get_colorant(color_string::String)
    try
        c = parse(Colors.Colorant, color_string)
        rgb = Colors.hex(c, :RRGGBB)
        return "FF" * rgb
    catch
        return nothing
    end
end
get_color(s::Symbol)::String = get_color(String(s))
function get_color(str::String)::String
    if occursin(r"^[0-9A-F]{8}$"i, str) # is a valid 8 digit hexadecimal color
        return uppercase(str)
    end
    s = replace(lowercase(str), "grey" => "gray")
    c = get_colorant(s)
    if isnothing(c)
        throw(XLSXError("Invalid color specified: $s. Either give a valid color name (from Colors.jl) or an 8-digit rgb color in the form AARRGGBB"))
    end
    return c
end

# This is Excel's 64-entry legacy indexed color palette
const INDEXED_PALETTE = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"
]

# Order required by the `theme` attribute on OOXML <color> elements.
# NOTE: this is NOT the document order of <a:clrScheme> (which lists dk1, lt1, dk2, lt2, ...).
# Excel swaps the dk/lt pairs when assigning indices: theme="0" means lt1, theme="1" means dk1,
# theme="2" means lt2, theme="3" means dk2. This is a long-documented OOXML quirk - see e.g.
# https://github.com/SheetJS/sheetjs/issues/389 - and is confirmed by every implementation that
# actually round-trips themes (e.g. ClosedXML maps Light1Color -> Background1 (index 0) and
# Dark1Color -> Text1 (index 1) in XLWorkbook_Load.cs).
const THEME_COLOR_ORDER = [
    "lt1", "dk1", "lt2", "dk2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink"
]

# Pull the RGB hex value out of a <a:dk1>/<a:lt1>/.../<a:accent6> element, which wraps
# either <a:srgbClr val="RRGGBB"/> or <a:sysClr val="windowText" lastClr="RRGGBB"/>.
function _theme_color_value(node::XML.Node)::Union{String,Nothing}
    for c in xml_elements(node)
        if localname(c) == "srgbClr" && haskey(c, "val")
            return uppercase(c["val"])
        elseif localname(c) == "sysClr" && haskey(c, "lastClr")
            return uppercase(c["lastClr"])
        end
    end
    return nothing
end

"""
    _themeelements_element(wb::Workbook) -> XML.Node

The theme's `<a:themeElements>`. Throws if the workbook's theme is malformed.
Shared by the colour and font scheme lookups.
"""
function _themeelements_element(wb::Workbook)::XML.Node
    xroot = xml_root_element(theme_xmlroot(wb))
    els = xml_elements(xroot)

    theme_els_idx = findfirst(c -> localname(c) == "themeElements", els)
    isnothing(theme_els_idx) &&
        throw(XLSXError("Malformed theme: no `themeElements` found in theme1.xml."))
    return els[theme_els_idx]
end

function _clrscheme_element(wb::Workbook)::XML.Node
    theme_els = _themeelements_element(wb)
    els = xml_elements(theme_els)

    clrscheme_idx = findfirst(c -> localname(c) == "clrScheme", els)
    isnothing(clrscheme_idx) &&
        throw(XLSXError("Malformed theme: no `clrScheme` found in theme1.xml."))
    return els[clrscheme_idx]
end

# Unlike `clrScheme`, a missing `fontScheme` is not fatal: nothing downstream
# needs a font the way `resolveColor` needs a palette, and `get_theme_fonts`
# reads an absent scheme as "no theme fonts" rather than a broken workbook.
_fontscheme_element(wb::Workbook) =
    first_element_with_tag(_themeelements_element(wb), "fontScheme")

# Read and cache the 12 theme colors for a workbook, in OOXML theme-index order
# (see `THEME_COLOR_ORDER`). Reads the actual `xl/theme/theme1.xml` clrScheme, so this
# reflects whatever theme the workbook was actually saved with - not just the Excel default.
function get_theme_colors(wb::Workbook)::Vector{String}
    if wb.theme_colors === nothing
        lookup = Dict{String,String}()
        for c in xml_elements(_clrscheme_element(wb))
            val = _theme_color_value(c)
            isnothing(val) || (lookup[localname(c)] = val)
        end
        wb.theme_colors = String[]
        for name in THEME_COLOR_ORDER
            if haskey(lookup, name)
                push!(wb.theme_colors, lookup[name])
            else
                @warn "Theme color $name missing in theme1.xml; using 000000 fallback"
                push!(wb.theme_colors, "000000")
            end
        end
    end
    return wb.theme_colors
end

# Excel tint algorithm
@inline function apply_tint(channel::UInt8, tint::Float64)::UInt8
    c = Float64(channel)
    if tint > 0.0
        c = c + (255.0 - c) * tint
    else
        c = c * (1.0 + tint)
    end
    return UInt8(clamp(round(Int, c), 0, 255))
end

# Convert theme + tint to RGB, using the workbook's actual theme colors.
function resolve_theme_color(wb::Workbook, theme_index::Int, tint::Float64)
    colors = get_theme_colors(wb)
    (theme_index < 0 || theme_index >= length(colors)) && throw(XLSXError("Invalid theme color index: $theme_index (expected 0..$(length(colors)-1))"))
    base = parse(UInt32, colors[theme_index + 1]; base=16)

    r = apply_tint(UInt8(base >> 16), tint)
    g = apply_tint(UInt8((base >> 8) & 0xFF), tint)
    b = apply_tint(UInt8(base & 0xFF), tint)

    buf = IOBuffer()
    print(buf, "FF")
    print(buf, uppercase(string(r, base=16, pad=2)))
    print(buf, uppercase(string(g, base=16, pad=2)))
    print(buf, uppercase(string(b, base=16, pad=2)))
    return String(take!(buf))
end

# get theme document (xl/theme/theme1.xml) for workbook.
# Unlike styles.xml, the theme part lives in the DrawingML namespace (NS_A), not the
# spreadsheet namespace, since themes are shared across Excel/Word/PowerPoint.
function theme_xmlroot(workbook::Workbook)
    if workbook.theme_xroot === nothing
        THEME_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        if has_relationship_by_type(workbook, THEME_RELATIONSHIP_TYPE)
            theme_target = get_relationship_target_by_type("xl", workbook, THEME_RELATIONSHIP_TYPE)
            theme_root = xmlroot(get_xlsxfile(workbook), theme_target)

            if get_default_namespace(xml_root_element(theme_root)) != NS_A
                throw(XLSXError("Unsupported theme XML namespace $(get_default_namespace(xml_root_element(theme_root)))."))
            end
            localname(xml_root_element(theme_root)) != "theme" && throw(XLSXError("Malformed package. Expected root node named `theme` in `theme1.xml`."))
            workbook.theme_xroot = theme_root
        else
            throw(XLSXError("Theme not found for this workbook."))
        end
    end
    return workbook.theme_xroot
end

"""
    resolveColor(ws::Worksheet, atts::AbstractDict; prefix::AbstractString="") -> String
    resolveColor(xl::XLSXFile, atts::AbstractDict; prefix::AbstractString="") -> String
    resolveColor(wb::Workbook, atts::AbstractDict; prefix::AbstractString="") -> String

Resolve a raw OOXML color attribute dictionary - of the kind returned in the `color` entry
of `getFont`, in each side's dictionary from `getBorder`, or in the `patternFill` dictionary
from `getFill` - to the "FFRRGGBB" color Excel would actually render in the worksheet, given 
the workbook's current theme.

`atts` can encode a color in one of four ways. `resolveColor` checks for them in this order:
- `rgb`     : an explicit 8-digit hex color. Returned unchanged.
- `theme`   : an index into the workbook's theme palette (`xl/theme/theme1.xml`), optionally
              adjusted by `tint`. This reads the workbook's actual theme colors, so it
              reflects custom themes correctly rather than assuming the Excel default.
- `indexed` : an index into the legacy 56-color indexed palette.
- `auto`    : Excel's automatic color, resolved to black ("FF000000").

If none of these keys is present, `resolveColor` returns "FF000000".

`getFont`'s `color` entry and each side of `getBorder`'s `border` dictionary use plain,
unprefixed keys, so the default `prefix=""` is correct for those. `getFill`'s `patternFill`
dictionary instead prefixes its two colors as `fgrgb`/`fgtheme`/`fgtint`/`fgindexed`/`fgauto`
and `bgrgb`/`bgtheme`/`bgtint`/`bgindexed`/`bgauto` (see `getFill`'s docstring) - pass
`prefix="fg"` or `prefix="bg"` to pick which one to resolve.

# Examples:
```julia
julia> resolveColor(sh, getFont(sh, "A1").font["color"])

julia> resolveColor(sh, getBorder(sh, "D4").border["top"])

julia> resolveColor(sh, getFill(sh, "D17").fill["patternFill"]; prefix="fg")

```
"""
function resolveColor end
resolveColor(ws::Worksheet, atts::AbstractDict; prefix::AbstractString="") = resolveColor(get_workbook(ws), atts; prefix)
resolveColor(xl::XLSXFile, atts::AbstractDict; prefix::AbstractString="") = resolveColor(get_workbook(xl), atts; prefix)
function resolveColor(wb::Workbook, atts::AbstractDict; prefix::AbstractString="")::String
    if haskey(atts, prefix*"rgb")
        raw = uppercase(strip(atts[prefix*"rgb"]))
        if length(raw) == 6
            return "FF" * raw
        elseif length(raw) == 8
            return raw
        else
            throw(XLSXError("Invalid rgb color format: $raw"))
        end

    elseif haskey(atts, prefix*"theme")
        rawtheme = strip(atts[prefix*"theme"])
        theme = try
            parse(Int, rawtheme)
        catch
            throw(XLSXError("Invalid theme index: $rawtheme"))
        end

        tint = 0.0
        if haskey(atts, prefix*"tint")
            rawt = strip(atts[prefix*"tint"])
            tint = try
                parse(Float64, rawt)
            catch
                throw(XLSXError("Invalid tint value: $rawt. Must be between -1.0 and 1.0."))
            end
            if !isfinite(tint)
                throw(XLSXError("Invalid tint value: $rawt. Must be between -1.0 and 1.0."))
            end
            tint = clamp(tint, -1.0, 1.0)
        end

        return resolve_theme_color(wb, theme, tint)

    elseif haskey(atts, prefix*"indexed")
        rawidx = strip(atts[prefix*"indexed"])
        idx = try
            parse(Int, rawidx)
        catch
            throw(XLSXError("Invalid indexed color index: $rawidx. Must be an integer between 0 and $(length(INDEXED_PALETTE)-1)."))
        end
        if idx < 0 || idx >= length(INDEXED_PALETTE)
            throw(XLSXError("Invalid indexed color index: $idx. Must be between 0 and $(length(INDEXED_PALETTE)-1)."))
        end
        return "FF" * INDEXED_PALETTE[idx+1]

    elseif haskey(atts, prefix*"auto")
        return "FF000000"
    else
        return "FF000000"
    end
end

"""
The theme's `clrScheme` keyed by DrawingML name, as `<a:schemeClr val="..."/>`
refers to it.

Distinct from [`get_theme_colors`](@ref), which is index-ordered for the
spreadsheet `<color theme="N"/>` attribute and applies Excel's documented dk/lt
index swap. Names are not swapped. `tx1`/`dk1` and `bg1`/`lt1` are aliases in
DrawingML, so both keys are populated.
"""
function get_theme_color_map(wb::Workbook)::Dict{String,String}
    if wb.theme_color_map === nothing
        m = Dict{String,String}()
        for c in xml_elements(_clrscheme_element(wb))
            val = _theme_color_value(c)
            isnothing(val) || (m[localname(c)] = val)
        end
        haskey(m, "dk1") && (m["tx1"] = m["dk1"])
        haskey(m, "lt1") && (m["bg1"] = m["lt1"])
        haskey(m, "dk2") && (m["tx2"] = m["dk2"])
        haskey(m, "lt2") && (m["bg2"] = m["lt2"])
        wb.theme_color_map = m
    end
    return wb.theme_color_map
end


"""
    get_theme_fonts(wb::Workbook) -> Dict{String,String}

The workbook theme's font scheme, keyed by the reference tokens as they appear
in DrawingML: `"+mj-lt"`, `"+mn-lt"`, `"+mj-ea"`, `"+mn-ea"`, `"+mj-cs"`,
`"+mn-cs"` (major/minor × latin/east-asian/complex-script). Cached on
`Workbook.theme_font_map`.

Empty typefaces are omitted rather than stored as `""`. Excel writes
`<a:ea typeface=""/>` to mean "no specific east-asian font, fall back", and a
missing key expresses that better than an empty string a caller has to test for.

Script-specific `<a:font script="Jpan" .../>` entries under each font are not
modelled; they are in the theme node if ever needed.

Distinct from `get_theme_color_map` only in what it reads — same theme part,
same caching, same name-keyed contract.
"""
function get_theme_fonts(wb::Workbook)::Dict{String,String}
    if wb.theme_font_map === nothing
        m = Dict{String,String}()
        for (tag, prefix) in (("majorFont", "+mj"), ("minorFont", "+mn"))
            group = first_element_with_tag(_fontscheme_element(wb), tag)
            for (child, suffix) in (("latin", "lt"), ("ea", "ea"), ("cs", "cs"))
                tf = _attr(first_element_with_tag(group, child), "typeface")
                isnothing(tf) || (m["$prefix-$suffix"] = tf)
            end
        end
        wb.theme_font_map = m
    end
    return wb.theme_font_map
end

"""
    resolve_theme_font(wb::Workbook, typeface) -> Union{Nothing,String}

Resolve a DrawingML typeface reference. `"+mn-lt"` and friends look up the
theme; any other string is a literal font name and passes through unchanged;
`nothing` (the attribute was absent) stays `nothing`, meaning "inherit".

Returns `nothing` for a theme reference the theme doesn't define, which is not
the same as the caller's font being unset — check `startswith(tf, '+')` first
if the distinction matters.
"""
resolve_theme_font(::Workbook, ::Nothing) = nothing
resolve_theme_font(wb::Workbook, typeface::AbstractString) =
    startswith(typeface, '+') ? get(get_theme_fonts(wb), typeface, nothing) : typeface

"""
Apply the DrawingML colour transforms to an "RRGGBB" hex string, in document
order, returning the transformed hex and the resulting alpha.

`lumMod`, `lumOff` and `satMod` operate in HSL, which is why they are applied
via Colors.jl rather than with the spreadsheet `apply_tint` algorithm: Excel's
`<color tint="...">` and DrawingML's transforms are different operations and
give different results.
"""
function apply_drawingml_transforms(hex::AbstractString,
                                    transforms::Vector{Pair{Symbol,Int}})
    c = Colors.RGB{Float64}(parse(Colors.Colorant, "#" * hex))
    alpha = 1.0

    for (kind, v) in transforms
        p = _pct(v)
        if kind === :alpha
            alpha = p
        elseif kind === :lumMod || kind === :lumOff || kind === :satMod
            h = convert(Colors.HSL{Float64}, c)
            l, s = h.l, h.s
            kind === :lumMod && (l *= p)
            kind === :lumOff && (l += p)
            kind === :satMod && (s *= p)
            c = convert(Colors.RGB{Float64},
                        Colors.HSL{Float64}(h.h, clamp(s, 0.0, 1.0), clamp(l, 0.0, 1.0)))
        elseif kind === :shade
            c = _linear_map(c, x -> x * p)
        elseif kind === :tint
            c = _linear_map(c, x -> x * p + (1.0 - p))
        end
        # Unhandled transforms (hueMod, red, green, blue, gamma, inv, gray,
        # comp) are rare in chart parts and are left as no-ops rather than
        # silently applied wrongly. The raw element is preserved regardless.
    end

    return Colors.hex(Colors.RGB{Float64}(clamp(c.r, 0, 1),
                                          clamp(c.g, 0, 1),
                                          clamp(c.b, 0, 1)), :RRGGBB), alpha
end

@inline _srgb_to_linear(x) = x <= 0.04045 ? x / 12.92 : ((x + 0.055) / 1.055)^2.4
@inline _linear_to_srgb(x) = x <= 0.0031308 ? x * 12.92 : 1.055 * x^(1/2.4) - 0.055

function _linear_map(c::Colors.RGB{Float64}, f)
    r = _linear_to_srgb(clamp(f(_srgb_to_linear(c.r)), 0.0, 1.0))
    g = _linear_to_srgb(clamp(f(_srgb_to_linear(c.g)), 0.0, 1.0))
    b = _linear_to_srgb(clamp(f(_srgb_to_linear(c.b)), 0.0, 1.0))
    return Colors.RGB{Float64}(r, g, b)
end