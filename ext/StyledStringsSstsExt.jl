module StyledStringsSstsExt 

using XLSX
using StyledStrings

# Import StyledStrings functions we need
import StyledStrings: AnnotatedString, Face, FACES, SimpleColor

# Import the function we're overriding
import XLSX: setdata!, ssToRuns, richTextString

# Colors.jl is already a dependency of XLSX
import Colors: RGB

# Import types we need
using XLSX: RichTextString, RichTextRun

include("ssExtTypes.jl")
include("ssExtHelpers.jl")

setdata!(sheet::Worksheet, ref::CellRef, ss::AnnotatedString{T}) where T = setdata!(sheet::Worksheet, ref::CellRef,richTextString(ssToRuns(ss)))
function ssToRuns(ss::AnnotatedString{T}) where T
    anns = annotations(ss)
    bounds = compute_bounds(ss.string, anns)
    segments = compute_segments(ss.string, bounds)

    runs = RichTextRun[]

    for seg in segments
        # 1. Collect raw annotation values for this segment
        raw_vals = active_faces(seg, anns)
        # 2. Convert each annotation value into a real Face
        faces = Face[]
        for v in raw_vals
            f = face_from_annotation(v)
            if f !== nothing
                push!(faces, f)
            end
        end

        # 3. Merge all Faces using StyledStrings.merge
        merged = isempty(faces) ? Face() :
                 reduce(StyledStrings.merge, faces; init=Face())


        # 4. Convert merged Face into Excel attributes
        atts = face_to_excel_atts(merged)

        # 5. Emit the run
        push!(runs, RichTextRun(ss.string[seg], atts))
    end

    return runs
end
function compute_bounds(str, anns)
    pts = Int[]
    push!(pts, firstindex(str))
    push!(pts, lastindex(str) + 1)  # half-open end

    for ann in anns
        r = ann.region
        push!(pts, first(r))
        push!(pts, nextind(str, last(r)))
    end

    return sort(unique(pts))
end

function compute_segments(str::String, bounds::Vector{Int})
    segs = UnitRange{Int}[]
    for i in 1:length(bounds)-1
        a = bounds[i]
        b = prevind(str, bounds[i+1])
        push!(segs, a:b)
    end

    return segs
end

function face_from_annotation(val)

    # Already a Face
    if val isa Face
        return val

        # SimpleColor wrapper
    elseif val isa SimpleColor
        return Face(foreground = resolve_color(val.value))

    # Raw RGB
    elseif val isa RGB
        return Face(foreground = val)

    # Symbol: could be a semantic face OR a color name
    elseif val isa Symbol
        if haskey(FACES, val)
            # semantic face like :warning, :error, :info
            return FACES[val]

        elseif val === :bold
            return Face(weight = :bold)

        elseif val === :error
            return Face(fgcolor = :error)

        elseif val === :italic
            return Face(slant = :italic)

        elseif val === :underline
            return Face(underline = true)

        elseif val === :strikethrough
            return Face(strikethrough = true)
        else
            # treat as color name
            return Face(foreground = resolve_color(val))
        end

    elseif val isa Pair
        println(val)
        key, v = val
        f = Face()

        if key === :foreground
            f = Face(foreground = resolve_color(v))

        elseif key === :background
            f = Face(background = resolve_color(v))

        elseif key === :color
            f = Face(foreground = resolve_color(v))

#            elseif key === :link
#                f = Face(link = v)

#            elseif key === :bold
#                f = Face(weight = :bold)

#            elseif key === :italic
#                f = Face(slant = :italic)

#            elseif key === :underline
#                f = Face(underline = true)
        end

        return f

    # Underline tuple or other structured annotation
#    elseif val isa Tuple
#        # handled in underline normalization later
#        return Face(underline = val)

    else
        return nothing
    end
end

function active_faces(seg, anns)
    faces = Face[]
    for ann in anns
        if !isempty(intersect(seg, ann.region))
            f = face_from_annotation(ann.value)
            if f !== nothing
                push!(faces, f)
            end
        end
    end
    return faces
end
function merge_two_faces(f1::Face, f2::Face)
    inh = f2.inherit

    # Determine base style
    base =
        inh === nothing      ? Face() :      # inherit nothing
        isempty(inh)         ? f1     :      # inherit everything
                               Face(         # inherit only selected fields
                                   font          = :font          in inh ? f1.font          : nothing,
                                   height        = :height        in inh ? f1.height        : nothing,
                                   weight        = :weight        in inh ? f1.weight        : nothing,
                                   slant         = :slant         in inh ? f1.slant         : nothing,
                                   foreground    = :foreground    in inh ? f1.foreground    : nothing,
                                   background    = :background    in inh ? f1.background    : nothing,
                                   underline     = :underline     in inh ? f1.underline     : false,
                                   strikethrough = :strikethrough in inh ? f1.strikethrough : false,
                                   inverse       = :inverse       in inh ? f1.inverse       : false,
                                   inherit       = Symbol[]
                               )

    # Overlay f2 on top of base
    return Face(
        font           = coalesce(f2.font,           base.font),
        height         = coalesce(f2.height,         base.height),
        weight         = coalesce(f2.weight,         base.weight),
        slant          = coalesce(f2.slant,          base.slant),
        foreground     = coalesce(f2.foreground,     base.foreground),
        background     = coalesce(f2.background,     base.background),
        underline      = boolmerge(f2.underline,      base.underline),
        strikethrough  = boolmerge(f2.strikethrough,  base.strikethrough),
        inverse        = boolmerge(f2.inverse,        base.inverse),
        inherit        = Symbol[]   # merged faces always inherit fully forward
    )
end
function merge_faces(faces::Vector{Face})
    merged = Face()
    for f in faces
        merged = merge_two_faces(merged, f)
    end
    return merged
end

const FONT_FAMILY_MAP = Dict(
    :mono  => "Consolas",
    :serif => "Times New Roman",
    :sans  => "Calibri"
)

function rgb_to_argb(rgb::RGB)
    r = round(Int, rgb.r * 255)
    g = round(Int, rgb.g * 255)
    b = round(Int, rgb.b * 255)
    return "FF" * hex(r) * hex(g) * hex(b)

end

function face_to_excel_atts(face)
    d = Dict{Symbol,Any}()

# available fields: `font`, `height`, `weight`, `slant`, `foreground`, `background`, `underline`, `strikethrough`, `inverse`, `inherit`

    # foreground → Excel color
    if face.foreground !== nothing
        c= ss_unwrap(face.foreground)
        fg=resolve_color(c)
        println(fg)
        if fg isa RGB
            # c is an actual RGB(r,g,b) object
            d[:color] = rgb_to_argb(fg)
        elseif fg isa Symbol || fg isa String
            #treat as a color from Colors.jl
            d[:color] = XLSX.get_color(fg)
        else
            println("unreachable reached")
            error()

        end
    end

#=
    # background (Excel doesn't support background in sharedStrings)
    # but you may want to record it anyway
    if face.background !== nothing
        d[:background] = ss_unwrap(face.background)
    end
=#

    if face.height !== nothing
        d[:size] = resolve_height(ss_unwrap(face.height), 11)
    end

    if face.font !== nothing
        f = ss_unwrap(face.font)
        if f isa Symbol
            d[:name] = get(FONT_FAMILY_MAP, f, String(f))

        else
            d[:name] = String(f)
        end
    end

    if face.weight !== nothing
        w = ss_unwrap(face.weight)
        if w isa Symbol
            d[:bold] = w in (:medium, :semibold, :bold, :extrabold, :black)
        elseif w isa Integer
            d[:bold] = w >= 600
        end
    end

    # booleans
    sl = ss_unwrap(face.slant)
    if sl === :italic
        d[:italic] = true
    end

    # no ability in XLSX to handle underline color as a separate attribute
    if !isnothing(face.underline)
        d[:under] = ssUnderlineToRtf(normalize_ssUnderline(ss_unwrap(face.underline)))
    end
    
    if face.strikethrough == true
        d[:strike] = true
    end

#=
    # vertical alignment (StyledStrings uses superscript/subscript)
    if face.superscript
        d[:vertAlign] = "superscript"
    elseif face.subscript
        d[:vertAlign] = "subscript"
    end
=#

    return d
end

end # module
