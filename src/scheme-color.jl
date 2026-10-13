# ===========================================================================
# Theme colors
#
# `SchemeColor` is shared by charts, which write it as a DrawingML `a:schemeClr`,
# and by cell formatting, conditional formats and rich text, which write it as a
# spreadsheet `<color theme="N" tint="t"/>` (see `_color_atts` in xlsx-colors.jl).
# ===========================================================================

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

A theme color: a token naming a slot in the workbook's color scheme, plus any
transforms applied to it. Because it refers to the theme rather than to a fixed
RGB value, it changes when the workbook's theme does, as Excel's own theme
colors do.

`token` is one of `:bg1` (background), `:tx1` (text), `:bg2`, `:tx2` (likewise),
`:accent1` … `:accent6`, `:hlink` (hyperlink), `:folHlink` (visited-hyperlink),
or the aliases `:lt1` (light), `:dk1` (dark), `:lt2`, `:dk2`.

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

# Cells, conditional formats and rich text

Charts accept any transforms. A cell color, though, can only be a theme color
made lighter or darker by a `tint`, so outside charts a `SchemeColor` may have:

- no transforms - the theme color itself;
- `lumMod = m` alone, with `m` from 0 to 100 - darker by `100 - m`%;
- `lumMod = m, lumOff = 100 - m` - lighter by `100 - m`%.

These are the shades Excel's own palette offers ("Darker 25%" is `lumMod = 75`,
"Lighter 40%" is `lumMod = 60, lumOff = 40`). An `alpha` transform is ignored,
because Excel ignores transparency in cell colors. Any other transform is an
error there; give an RGB color instead.

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
