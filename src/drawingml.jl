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
function color_element(node::Union{Nothing,XML.Node})::Union{Nothing,XML.Node}
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
    el = localname(node) in DML_COLOR_TAGS ? node : color_element(node)
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
function fill_element(node::Union{Nothing,XML.Node})::Union{Nothing,XML.Node}
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
    el = haskey(DML_FILL_TAGS, localname(node)) ? node : fill_element(node)
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


"""
    parse_drawing_line(wb, node) -> Union{Nothing,DrawingLine}

Parse the outline held by `node`, or `nothing` where it holds none. Pass the
`a:ln` element itself or its parent.
"""
function parse_drawing_line(wb::Workbook, node::Union{Nothing,XML.Node})::Union{Nothing,DrawingLine}
    isnothing(node) && return nothing
    el = localname(node) == "ln" ? node : first_element_with_tag(node, "ln")
    isnothing(el) && return nothing

    dash = first_element_with_tag(el, "prstDash")

    return DrawingLine(
        parse_drawing_fill(wb, el),
        _attr_emu(el, "w"),          # points, converted at parse time
        _attr(dash, "val"),
        _attr(el, "cap"),
        _attr(el, "cmpd"),
        el,
    )
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

# Inheritance. Every field on DrawingRunProps is Union{Nothing,T}: absent means
# "inherit", not "default", because a run's a:rPr sets only what differs from
# what it inherits. The cascade, innermost first:
#
#   1. the run's own a:rPr
#   2. the paragraph's a:pPr/a:defRPr
#   3. the body's a:lstStyle/a:lvl<n>pPr/a:defRPr (n from a:pPr/@lvl)
#   4. the c:txPr of the enclosing element (axis, series, data labels, legend)
#   5. c:chartSpace/c:txPr
#   6. the theme font scheme for +mj-*/+mn-* typefaces — get_theme_fonts —
#      and Excel's built-in defaults for everything else
#
# Levels 4 and 5 need a parent chain that doesn't exist until stage 3, and the
# rules differ by site (a data label does not inherit like an axis title), so
# the resolver is stage 4 work. Until then nothing user-facing should promise a
# resolved value: parsers return what is written, and `default_run_props`
# returns the nearest as-written properties, not the effective ones.

# =============================================================================
# DrawingML text: CT_TextBody
#
# The same content model serves `c:txPr` (formatting-only: axes, data labels,
# legend), `c:rich` (text with content: titles), and the cx: equivalent. One
# set of parsers; the caller says which child tag to look for.
#
# Two shapes to keep in mind, because they read differently:
#
#   txPr  — no runs at all. Font lives in `a:pPr/a:defRPr`, and Excel writes an
#           empty `a:endParaRPr` after it. `default_run_props` resolves this.
#   rich  — one or more `a:r`, each with its own `a:rPr` setting only what
#           differs from `defRPr`. May be uniform or mixed; see `is_uniform`.
#
# As with `prstDash`, several things that look like attributes are elements:
# autofit (`a:noAutofit`/`a:normAutofit`/`a:spAutoFit`), line spacing
# (`a:lnSpc/a:spcPct`), and the typefaces (`a:latin`, `a:ea`, `a:cs`).
#
# Write-back note for stage 4: `a:rPr` is an xsd:sequence — ln, fill,
# effectLst, highlight, uLn*, uFill*, latin, ea, cs, sym, hlink*, rtl, extLst.
# Splice into `raw`; never regenerate.
# =============================================================================

# ---------------------------------------------------------------------------
# Attribute readers
# ---------------------------------------------------------------------------

const EMU_PER_POINT = 12700

_attr_int(node, key) = (s = _attr(node, key); s === nothing ? nothing : tryparse(Int, s))

_attr_bool(node, key) =
    (s = _attr(node, key); s === nothing ? nothing : (s == "1" || s == "true" || s == "on"))

# 1/100 pt -> pt   (sz, kern, spc)
_attr_pt(node, key) = (n = _attr_int(node, key); n === nothing ? nothing : n / 100)

# EMU -> pt        (marL, marR, indent, insets)
_attr_emu(node, key) = (n = _attr_int(node, key); n === nothing ? nothing : n / EMU_PER_POINT)

# 1/60000 deg -> deg  (rot)
_attr_deg(node, key) = (n = _attr_int(node, key); n === nothing ? nothing : n / 60_000)

# As `_attr_pct`, but returns `nothing` rather than a default when the
# attribute is absent — DrawingML text needs absent and explicit to stay
# distinct. Same units: thousandths of a percent in, fraction out (60000 -> 0.6).
_attr_pct_opt(node, key) = (n = _attr_int(node, key); n === nothing ? nothing : _pct(n))


"""
    _attr_spacing(parent, tag) -> Union{Nothing,Tuple{Symbol,Float64}}

`a:lnSpc`, `a:spcBef` and `a:spcAft` each wrap either `a:spcPct` (1/1000 %) or
`a:spcPts` (1/100 pt). Returns `(:pct, 150.0)` or `(:pts, 12.0)`.
"""
function _attr_spacing(parent::XML.Node, tag::AbstractString)
    el = first_element_with_tag(parent, tag)
    el === nothing && return nothing
    pct = first_element_with_tag(el, "spcPct")
    pct === nothing || return (:frac, _attr_pct_opt(pct, "val"))
    pts = first_element_with_tag(el, "spcPts")
    pts === nothing || return (:pts, _attr_pt(pts, "val"))
    return nothing
end

# ---------------------------------------------------------------------------
# a:rPr / a:defRPr / a:endParaRPr
# ---------------------------------------------------------------------------

"""
    parse_drawing_run_props(wb, node; tag="rPr") -> Union{Nothing,DrawingRunProps}

Parse a run-properties element. `node` may be the element itself or its parent;
`tag` selects `rPr`, `defRPr` or `endParaRPr` when searching a parent.

Fields left `nothing` are absent from the file, which means "inherit" — not
"default". Resolving the cascade is `effective_run_props`' job, not this one's.
"""
function parse_drawing_run_props(wb::Workbook, node::XML.Node; tag::AbstractString="rPr")
    el = localname(node) == tag ? node : first_element_with_tag(node, tag)
    el === nothing && return nothing

    return DrawingRunProps(
        _attr(el, "lang"),
        _attr_pt(el, "sz"),
        _attr_bool(el, "b"),
        _attr_bool(el, "i"),
        _attr(el, "u"),          # none | sng | dbl | heavy | dotted | ...
        _attr(el, "strike"),     # noStrike | sngStrike | dblStrike
        _attr(el, "cap"),        # none | small | all
        _attr_pct_opt(el, "baseline"),
        _attr_pt(el, "kern"),
        _attr_pt(el, "spc"),
        parse_drawing_fill(wb, el),
        parse_drawing_line(wb, el),
        _attr(first_element_with_tag(el, "latin"), "typeface"),
        _attr(first_element_with_tag(el, "ea"), "typeface"),
        _attr(first_element_with_tag(el, "cs"), "typeface"),
        el,
    )
end

# ---------------------------------------------------------------------------
# a:pPr
# ---------------------------------------------------------------------------

"""
    parse_drawing_paragraph_props(wb, node; tag="pPr") -> Union{Nothing,DrawingParaProps}

Parse paragraph properties. The nested `a:defRPr` is parsed into a full
`DrawingRunProps` — in a chart `txPr` it is the only place the font appears.
"""
function parse_drawing_paragraph_props(wb::Workbook, node::XML.Node; tag::AbstractString="pPr")
    el = localname(node) == tag ? node : first_element_with_tag(node, tag)
    el === nothing && return nothing

    return DrawingParaProps(
        _attr(el, "algn"),       # l | ctr | r | just | justLow | dist | thaiDist
        _attr_int(el, "lvl"),
        _attr_emu(el, "marL"),
        _attr_emu(el, "marR"),
        _attr_emu(el, "indent"),
        _attr_bool(el, "rtl"),
        _attr_spacing(el, "lnSpc"),
        _attr_spacing(el, "spcBef"),
        _attr_spacing(el, "spcAft"),
        parse_drawing_run_props(wb, el; tag="defRPr"),
        el,
    )
end

# ---------------------------------------------------------------------------
# a:bodyPr
# ---------------------------------------------------------------------------

"""
    parse_drawing_body_props(node; tag="bodyPr") -> Union{Nothing,DrawingBodyProps}

Parse text-body properties. Autofit is an element, not an attribute:
`a:noAutofit` -> `:none`, `a:normAutofit` -> `:normal` (with `fontscale` and
`linespacereduction`), `a:spAutoFit` -> `:shape`. Absent means "inherit".
"""
function parse_drawing_body_props(node::XML.Node; tag::AbstractString="bodyPr")
    el = localname(node) == tag ? node : first_element_with_tag(node, tag)
    el === nothing && return nothing

    autofit = nothing
    fontscale = nothing
    lnspcred = nothing
    if first_element_with_tag(el, "noAutofit") !== nothing
        autofit = :none
    elseif (na = first_element_with_tag(el, "normAutofit")) !== nothing
        autofit = :normal
        fontscale = _attr_pct_opt(na, "fontScale")
        lnspcred = _attr_pct_opt(na, "lnSpcReduction")
    elseif first_element_with_tag(el, "spAutoFit") !== nothing
        autofit = :shape
    end

    return DrawingBodyProps(
        _attr_deg(el, "rot"),
        _attr(el, "vert"),          # horz | vert | vert270 | wordArtVert | ...
        _attr(el, "wrap"),          # none | square
        _attr(el, "anchor"),        # t | ctr | b | just | dist
        _attr_bool(el, "anchorCtr"),
        _attr_bool(el, "upright"),
        _attr_bool(el, "spcFirstLastPara"),
        _attr(el, "vertOverflow"),  # overflow | ellipsis | clip
        _attr(el, "horzOverflow"),  # overflow | clip
        _attr_emu(el, "lIns"),
        _attr_emu(el, "tIns"),
        _attr_emu(el, "rIns"),
        _attr_emu(el, "bIns"),
        autofit,
        fontscale,
        lnspcred,
        el,
    )
end

"""
    parse_drawing_shape_props(wb, node; tag="spPr") -> Union{Nothing,DrawingShapeProps}

Parse a shape-properties element. `node` may be the element itself or its
parent, and may be `nothing` so that optional lookups chain without guards:

    parse_drawing_shape_props(wb, series_node)                    # c:spPr
    parse_drawing_shape_props(wb, first_element_with_tag(x, "y")) # may be nothing

Returns `nothing` when there is no `spPr` — which for a series means "inherit
from the chart style", not "no fill".
"""
function parse_drawing_shape_props(wb::Workbook, node::Union{Nothing,XML.Node};
                                   tag::AbstractString="spPr")::Union{Nothing,DrawingShapeProps}
    isnothing(node) && return nothing
    el = localname(node) == tag ? node : first_element_with_tag(node, tag)
    isnothing(el) && return nothing

    effects = first_element_with_tag(el, "effectLst")
    isnothing(effects) && (effects = first_element_with_tag(el, "effectDag"))

    return DrawingShapeProps(
        parse_drawing_fill(wb, el),
        parse_drawing_line(wb, el),
        effects,
        _attr(el, "bwMode"),
        el,
    )
end

# ---------------------------------------------------------------------------
# Convenience predicates
#
# These exist because `isnothing(sp.fill)` and `sp.fill.kind == :none` are easy
# to conflate at a call site, and the difference is the whole point of the type.
# ---------------------------------------------------------------------------

"""
    has_fill(sp::DrawingShapeProps) -> Bool

Whether the shape sets a visible fill. `false` both when no fill is specified
(inherited) and when `<a:noFill/>` is written (deliberately transparent) — use
`sp.fill` directly to tell those apart.
"""
has_fill(sp::DrawingShapeProps) = !isnothing(sp.fill) && sp.fill.kind !== :none

"""
    has_line(sp::DrawingShapeProps) -> Bool

Whether the shape sets a visible outline. `false` when `a:ln` is absent, and
also when it contains `<a:noFill/>`, which is how Excel writes "no border".
"""
has_line(sp::DrawingShapeProps) =
    !isnothing(sp.line) && !isnothing(sp.line.fill) && sp.line.fill.kind !== :none

function Base.show(io::IO, sp::DrawingShapeProps)
    parts = String[]
    if isnothing(sp.fill)
        push!(parts, "fill inherited")
    else
        push!(parts, "fill $(sp.fill.kind)")
    end
    if isnothing(sp.line)
        push!(parts, "line inherited")
    elseif !isnothing(sp.line.fill) && sp.line.fill.kind === :none
        push!(parts, "no line")
    else
        push!(parts, isnothing(sp.line.width) ? "line" : "line $(sp.line.width)pt")
    end
    isnothing(sp.effects) || push!(parts, "effects")
    print(io, "DrawingShapeProps(", join(parts, ", "), ")")
end

# ---------------------------------------------------------------------------
# a:p and its runs
# ---------------------------------------------------------------------------

"""
    parse_drawing_paragraph(wb, p) -> DrawingParagraph

`a:r`, `a:br` and `a:fld` interleave in document order, so this walks children
rather than using `elements_with_tag`. Non-element children (whitespace in a
formatted part, comments) are skipped — Excel writes chart parts unindented,
but a part that has been through a formatter or hand-edited will have them.

`a:fld` (page numbers and the like) is kept as a run with its cached `a:t` text
and `kind == :fld`; anything richer stays in `raw`.
"""
function parse_drawing_paragraph(wb::Workbook, p::XML.Node)
    runs = DrawingRun[]
    for c in XML.children(p)
        XML.nodetype(c) === XML.Element || continue
        ln = localname(c)
        if ln == "r"
            push!(runs, DrawingRun(:run, child_text(c, "t"), parse_drawing_run_props(wb, c), c))
        elseif ln == "br"
            push!(runs, DrawingRun(:br, "\n", parse_drawing_run_props(wb, c), c))
        elseif ln == "fld"
            push!(runs, DrawingRun(:fld, child_text(c, "t"), parse_drawing_run_props(wb, c), c))
        end
    end

    return DrawingParagraph(
        parse_drawing_paragraph_props(wb, p),
        runs,
        parse_drawing_run_props(wb, p; tag="endParaRPr"),
        p,
    )
end

# ---------------------------------------------------------------------------
# CT_TextBody
# ---------------------------------------------------------------------------

"""
    parse_drawing_text(wb, node; tag="txPr") -> Union{Nothing,DrawingText}

Parse a DrawingML text body. `node` may be the text body itself (`c:txPr`,
`c:rich`, `cx:txPr`) or its parent, in which case `tag` selects the child.

    parse_drawing_text(wb, axis_node)                 # c:txPr
    parse_drawing_text(wb, tx_node; tag = "rich")     # title text

`a:lstStyle` is preserved as a node only: in chart parts it is either empty or
carries list-level defaults we do not model, and dropping it would change how
Excel renders inherited text.
"""
function parse_drawing_text(wb::Workbook, node::XML.Node; tag::AbstractString="txPr")
    el = if localname(node) == tag || _is_text_body(node)
        node
    else
        first_element_with_tag(node, tag)
    end
    el === nothing && return nothing

    paragraphs = [parse_drawing_paragraph(wb, p) for p in elements_with_tag(el, "p")]

    return DrawingText(
        parse_drawing_body_props(el),
        first_element_with_tag(el, "lstStyle"),
        paragraphs,
        el,
    )
end

_is_text_body(node::XML.Node) =
    first_element_with_tag(node, "bodyPr") !== nothing && localname(node) != "spPr"

# ---------------------------------------------------------------------------
# Reading the text back out
# ---------------------------------------------------------------------------

"""
    text_content(t::DrawingText) -> String

Concatenate run text, one line per paragraph. Empty for a formatting-only
`txPr`; the visible string for a `c:rich` title.

Excel splits runs on language and spellcheck boundaries, so a title typed as
one string may arrive as several runs with identical properties. This
reassembles it; editing does not — replacing whole text means collapsing to a
single run, preserving the first run's `rPr`.
"""
function text_content(t::DrawingText)
    io = IOBuffer()
    for (i, p) in enumerate(t.paragraphs)
        i > 1 && print(io, "\n")
        for r in p.runs
            r.text !== nothing && print(io, r.text)
        end
    end
    return String(take!(io))
end

text_runs(t::DrawingText) = [r for p in t.paragraphs for r in p.runs]

# Comparable identity for run properties, ignoring `lang` and `raw`: two runs
# differing only by spellcheck language are the same formatting.
#
# Two conventions shared with `Base.:(==)(::RichTextRun, ::RichTextRun)`:
# colours compare on their resolved value rather than how they were written
# (as that does via `get_color`), and an absent property is distinct from one
# set to a default (as that does by comparing key sets).
_color_key(c) = c === nothing ? nothing : (c.rgb, c.alpha)
_fill_key(f) = f === nothing ? nothing : (f.kind, _color_key(f.fgcolor), _color_key(f.bgcolor), f.preset)
_line_key(l) = l === nothing ? nothing : (_fill_key(l.fill), l.width, l.dash, l.cap, l.compound)

_props_key(::Nothing) = nothing
_props_key(p::DrawingRunProps) = (
    p.size, p.bold, p.italic, p.under, p.strike, p.caps,
    p.baseline, p.kern, p.spacing,
    _fill_key(p.fill), _line_key(p.line),
    p.latin, p.ea, p.cs,
)

"""
    is_uniform(t::DrawingText) -> Bool

Whether every run in `t` carries the same formatting, comparing only what is
written on each `a:rPr` (a run with no `rPr` falls back to its paragraph's
`defRPr`). Text with no runs is uniform by definition.

This compares *as written*, not as resolved — two runs reaching the same
appearance by different inheritance paths compare as different. That is the
conservative direction: it can report mixed where a full cascade would report
uniform, but never the reverse.

Unlike `==` on `RichTextRun`, this ignores run text entirely: the question is
whether formatting is consistent across the text, not whether the runs are
identical. Two runs reading differently but formatted alike are uniform.
"""
function is_uniform(t::DrawingText)
    key = nothing
    seen = false
    for p in t.paragraphs
        fallback = p.props === nothing ? nothing : p.props.defprops
        for r in p.runs
            k = _props_key(r.props === nothing ? fallback : r.props)
            if !seen
                key = k
                seen = true
            elseif k != key
                return false
            end
        end
    end
    return true
end

"""
    default_run_props(t::DrawingText) -> Union{Nothing,DrawingRunProps}

The run properties describing this text body's appearance as a whole: the
first paragraph's `a:defRPr` if present, else the first run's `a:rPr`, else the
first `a:endParaRPr`.

Returns `nothing` for text whose runs disagree — a mixed-format title has no
single answer, and returning the first run's properties would quietly describe
only its opening fragment. Use `first_run_props` if you want that value anyway,
or `is_uniform` to test first.
"""
function default_run_props(t::DrawingText)
    is_uniform(t) || return nothing
    return first_run_props(t)
end

"""
    first_run_props(t::DrawingText) -> Union{Nothing,DrawingRunProps}

The first run properties found, in `defRPr` -> `rPr` -> `endParaRPr` order,
regardless of whether later runs agree. For a formatting-only `txPr` this is
the whole story; for mixed text it describes only the first fragment.
"""
function first_run_props(t::DrawingText)
    for p in t.paragraphs
        p.props !== nothing && p.props.defprops !== nothing && return p.props.defprops
        !isempty(p.runs) && p.runs[1].props !== nothing && return p.runs[1].props
        p.endprops !== nothing && return p.endprops
    end
    return nothing
end

function Base.show(io::IO, t::DrawingText)
    s = text_content(t)
    np = length(t.paragraphs)
    nr = sum(length(p.runs) for p in t.paragraphs; init=0)
    if isempty(s)
        print(io, "DrawingText($np paragraph(s), formatting only)")
    else
        mixed = is_uniform(t) ? "" : ", mixed"
        print(io, "DrawingText(", repr(truncate_len(s, 40)), ", $np paragraph(s), $nr run(s)$mixed)")
    end
end