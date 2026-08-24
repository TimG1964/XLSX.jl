#
# drawingml.jl
#
# Shared DrawingML parsing: the `a:` namespace elements that appear throughout
# chart parts, drawings and themes.
#
# Colours first. Every fill, line and text run in DrawingML resolves to one of
# six colour elements, optionally modified by transforms. Parsing this once and
# reusing it is what keeps the chart formatting work tractable: `spPr` alone
# appears on the chart space, plot area, every series, every data point, every
# axis, the legend, the title, data labels, gridlines, trendlines and error
# bars.
#

# The six colour elements. All carry a `val` attribute except `scrgbClr`, which
# uses separate r/g/b percentage attributes.
const DML_COLOR_TAGS = ("srgbClr", "schemeClr", "sysClr", "prstClr", "hslClr", "scrgbClr")

const DML_COLOR_KIND = Dict(
    "srgbClr"  => :srgb,
    "schemeClr"=> :scheme,
    "sysClr"   => :sys,
    "prstClr"  => :prst,
    "hslClr"   => :hsl,
    "scrgbClr" => :scrgb,
)

# Transforms we resolve. Others (hueMod, red, green, blue, gamma, inv, gray,
# comp) are recorded but not applied: they are vanishingly rare in chart parts,
# and applying them wrongly would be worse than leaving the base colour.
const DML_APPLIED_TRANSFORMS = (:lumMod, :lumOff, :satMod, :shade, :tint, :alpha)

"""
The DrawingML colour element among the children of `node`, or `nothing`.

A fill, line or run-property element holds at most one colour child, so this
returns the first match rather than a list.
"""
function first_color_element(node::Union{Nothing,XML.Node})::Union{Nothing,XML.Node}
    isnothing(node) && return nothing
    for child in XML.eachelement(node)
        localname(child) in DML_COLOR_TAGS && return child
    end
    return nothing
end

"""
Resolve a colour reference to `"RRGGBB"`, before any transforms are applied.

`schemeClr` goes through the workbook theme; `sysClr` prefers the `lastClr`
attribute Excel writes alongside the system colour name, since the name itself
depends on OS settings; `prstClr` goes through the Colors.jl named colours.
"""
function resolve_color_base(wb::Workbook, node::XML.Node)::String
    tag = localname(node)
    val = get_attr(node, "val")

    if tag == "srgbClr"
        return uppercase(val)

    elseif tag == "schemeClr"
        m = get_theme_color_map(wb)
        haskey(m, val) && return uppercase(m[val])
        # phClr is a placeholder resolved by the style that references it, and
        # has no meaning on its own. Anything else is an unknown scheme name.
        return "000000"

    elseif tag == "sysClr"
        last = get_attr(node, "lastClr")
        isempty(last) || return uppercase(last)
        return val == "window" ? "FFFFFF" : "000000"

    elseif tag == "prstClr"
        # DrawingML preset names are camelCase X11, except that dark/light/medium
        # are abbreviated, which Colors.jl does not know.
        s = lowercase(val)
        s = startswith(s, "dk")  ? "dark"   * s[3:end] :
            startswith(s, "lt")  ? "light"  * s[3:end] :
            startswith(s, "med") ? "medium" * s[4:end] : s
        c = get_colorant(replace(s, "grey" => "gray"))
        return isnothing(c) ? "000000" : uppercase(c[3:end])

    elseif tag == "hslClr"
        h = _attr_pct(node, "hue", 0.0) * 360.0 / 100.0   # hue is in 1/60000 degree
        hd = tryparse(Int, get_attr(node, "hue"))
        hue = isnothing(hd) ? 0.0 : hd / 60_000.0
        s = _attr_pct(node, "sat", 0.0)
        l = _attr_pct(node, "lum", 0.0)
        c = convert(Colors.RGB{Float64}, Colors.HSL{Float64}(hue, s, l))
        return Colors.hex(c, :RRGGBB)

    elseif tag == "scrgbClr"
        c = Colors.RGB{Float64}(_attr_pct(node, "r", 0.0),
                                _attr_pct(node, "g", 0.0),
                                _attr_pct(node, "b", 0.0))
        return Colors.hex(c, :RRGGBB)
    end

    return "000000"
end

# Percentage attributes are in thousandths of a percent, as elsewhere in DrawingML.
@inline function _attr_pct(node::XML.Node, name::String, default::Float64)::Float64
    s = get_attr(node, name)
    isempty(s) && return default
    v = tryparse(Int, s)
    return isnothing(v) ? default : v / 100_000.0
end

"""
    parse_drawing_color(wb, node) -> Union{Nothing,DrawingColor}

Parse the DrawingML colour held by `node` - a fill, a line, a run property -
into a [`DrawingColor`](@ref), or `nothing` where it holds none.

Pass the colour element itself, or its parent: a parent is searched for its
colour child.
"""
function parse_drawing_color(wb::Workbook, node::Union{Nothing,XML.Node})::Union{Nothing,DrawingColor}
    isnothing(node) && return nothing
    el = localname(node) in DML_COLOR_TAGS ? node : first_color_element(node)
    isnothing(el) && return nothing

    kind = DML_COLOR_KIND[localname(el)]
    val  = get_attr(el, "val")
    base = resolve_color_base(wb, el)

    transforms = Pair{Symbol,Int}[]
    for child in XML.eachelement(el)
        v = tryparse(Int, get_attr(child, "val"))
        isnothing(v) && continue
        push!(transforms, Symbol(localname(child)) => v)
    end

    rgb, alpha = apply_drawingml_transforms(base, transforms)
    return DrawingColor(kind, val, transforms, rgb, alpha)
end

Base.show(io::IO, c::DrawingColor) =
    print(io, "XLSX.DrawingColor(", c.val, " -> #", c.rgb,
          c.alpha == 1.0 ? "" : ", alpha " * string(round(c.alpha; digits=3)), ")")

function Base.show(io::IO, ::MIME"text/plain", c::DrawingColor)
    println(io, "XLSX.DrawingColor ", c.kind, " ", repr(c.val))
    println(io, "  resolves to: #", c.rgb)
    c.alpha == 1.0 || println(io, "  alpha: ", round(c.alpha; digits=3))
    isempty(c.transforms) ||
        println(io, "  transforms: ",
                join(("$k=$(v/1000)%" for (k, v) in c.transforms), ", "))
end

const DML_FILL_TAGS = Dict(
    "noFill"   => :none,
    "solidFill"=> :solid,
    "gradFill" => :gradient,
    "pattFill" => :pattern,
    "blipFill" => :blip,
    "grpFill"  => :group,
)

"""
The fill element among the children of `node`, or `nothing` where there is
none. An `spPr` holds at most one.
"""
function first_fill_element(node::Union{Nothing,XML.Node})::Union{Nothing,XML.Node}
    isnothing(node) && return nothing
    for child in XML.eachelement(node)
        haskey(DML_FILL_TAGS, localname(child)) && return child
    end
    return nothing
end

"""
    parse_drawing_fill(wb, node) -> Union{Nothing,DrawingFill}

Parse the fill held by `node`, or `nothing` where it holds none. Pass the fill
element itself or its parent.
"""
function parse_drawing_fill(wb::Workbook, node::Union{Nothing,XML.Node})::Union{Nothing,DrawingFill}
    isnothing(node) && return nothing
    el = haskey(DML_FILL_TAGS, localname(node)) ? node : first_fill_element(node)
    isnothing(el) && return nothing
    kind = DML_FILL_TAGS[localname(el)]

    if kind === :solid
        return DrawingFill(kind, parse_drawing_color(wb, el), nothing, nothing, el)

    elseif kind === :pattern
        fg = parse_drawing_color(wb, first_element_with_tag(el, "fgClr"))
        bg = parse_drawing_color(wb, first_element_with_tag(el, "bgClr"))
        prst = get_attr(el, "prst")
        return DrawingFill(kind, fg, bg, isempty(prst) ? nothing : prst, el)
    end

    return DrawingFill(kind, nothing, nothing, nothing, el)
end

Base.show(io::IO, fl::DrawingFill) =
    print(io, "XLSX.DrawingFill(", fl.kind,
          isnothing(fl.fgcolor) ? "" : ", #" * fl.fgcolor.rgb,
          isnothing(fl.bgcolor) ? "" : " on #" * fl.bgcolor.rgb, ")")

function Base.show(io::IO, ::MIME"text/plain", fl::DrawingFill)
    println(io, "XLSX.DrawingFill ", fl.kind)
    isnothing(fl.preset)  || println(io, "  pattern: ", fl.preset)
    isnothing(fl.fgcolor) || println(io, "  foreground: #", fl.fgcolor.rgb)
    isnothing(fl.bgcolor) || println(io, "  background: #", fl.bgcolor.rgb)
    fl.kind in (:gradient, :blip, :group) &&
        println(io, "  (not modelled; preserved on write)")
end