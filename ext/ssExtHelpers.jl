const ANSI_RGB = Dict(
    :bright_red     => RGB(0xed/255, 0x33/255, 0x3b/255),
    :bright_green   => RGB(0x33/255, 0xd0/255, 0x79/255),
    :bright_yellow  => RGB(0xf6/255, 0xd2/255, 0x2c/255),
    :bright_blue    => RGB(0x35/255, 0x83/255, 0xe4/255),
    :bright_magenta => RGB(0xbf/255, 0x60/255, 0xca/255),
    :bright_cyan    => RGB(0x26/255, 0xc6/255, 0xda/255),
    :bright_white   => RGB(0xf6/255, 0xf5/255, 0xf4/255),
    :red            => RGB(0xa5/255, 0x1c/255, 0x2c/255),
    :yellow         => RGB(0xe5/255, 0xa5/255, 0x09/255),
    :green          => RGB(0x25/255, 0xa2/255, 0x68/255),
    :blue           => RGB(0x19/255, 0x5e/255, 0xb3/255),
    :magenta        => RGB(0x80/255, 0x3d/255, 0x9b/255),
    :cyan           => RGB(0x00/255, 0x97/255, 0xa7/255),
    :white          => RGB(0xdd/255, 0xdc/255, 0xd9/255),
    :black          => RGB(0x1c/255, 0x1a/255, 0x23/255),
)
const FACE_TABLE = StyledStrings.FACES.current[]

# Debug!
function dump_face(f::StyledStrings.Face)
    println("Face(")
    for name in fieldnames(StyledStrings.Face)
        val = getfield(f, name)
        println("  $name = $val")
    end
    println(")")
end

ss_unwrap(x) = hasproperty(x, :value) ? getproperty(x, :value) : x

boolmerge(child, parent) = child === nothing ? parent : child

hex(x) = uppercase(string(x, base=16, pad=2))

function resolve_height(height, base_size_pt)
    if height === nothing
        return base_size_pt
    elseif height isa Int
        return height / 10        # deci-pt → pt
    elseif height isa Float64
        return base_size_pt * height
    else
        error("Unexpected height type: $(typeof(height))")
    end
end

function normalize_ssUnderline(u)
    if u === nothing
        return ssUnderlineSpec(nothing, nothing)

    elseif u isa Bool
        return u ? ssUnderlineSpec(nothing, :straight) :
                   ssUnderlineSpec(nothing, nothing)

    elseif u isa StyledStrings.SimpleColor
        return ssUnderlineSpec(u.value, :straight)

    elseif u isa Symbol
        # Look up the symbol in the global FACES dictionary
        if haskey(StyledStrings.FACES, u)
            f = StyledStrings.FACES[u]
            return normalize_ssUnderline(f.underline)
        else
            # Interpret unknown symbols as underline color
            return ssUnderlineSpec(u, :straight)
        end

    elseif u isa Tuple
        color, style = u
        if color isa StyledStrings.SimpleColor
            return ssUnderlineSpec(color.value, style)
        elseif color === nothing
            return ssUnderlineSpec(nothing, style)
        else
            error("Invalid underline tuple: $u")
        end

    else
        error("Unknown underline type: $(typeof(u))")
    end
end
function ssUnderlineToRtf(spec::ssUnderlineSpec)
#    d = Dict{Symbol,Any}()

    if spec.style === nothing
        return nothing
    end

    # style
    under = spec.style === :double ? "double" : "single"
#    d[:underline] = excel_style

    # color
#    if spec.color !== nothing
#        d[:underline_color] = color_to_argb(spec.color)
#    end

    return under
end

function resolve_color(c)

    # 1. Already an RGB color
    if c isa RGB
        return c

    # 2. NamedTuple(r, g, b)
    elseif c isa NamedTuple
        return RGB(c.r/255, c.g/255, c.b/255)

    # 3. unwrap SimpleColor
    elseif c isa SimpleColor
        return resolve_color(c.value)

    # 4. Create RGB from a string (is this possible?)
    elseif c isa String
        println(c)
        hex = replace(c, "#" => "")
        r = parse(Int, hex[1:2], base=16) / 255
        g = parse(Int, hex[3:4], base=16) / 255
        b = parse(Int, hex[5:6], base=16) / 255
        return RGB(r, g, b)

    elseif c isa Symbol
        # 5. ANSI / bright color mapping
        if haskey(ANSI_RGB, c)
            return ANSI_RGB[c]
        end

        # 6. Look for a color from Colors.jl
        rgb = XLSX.get_colorant(c)
        isnothing(rgb) || return rgb

         # 7. If it's a semantic face, unwrap once
       if haskey(FACE_TABLE, c)
            fg = FACE_TABLE[c].foreground
            # If the face points to itself, stop
            if fg === c
                return nothing
            end
            return resolve_color(fg)
        end

        # 8. uh-oh!
        println("Unreachable reached!")
        error()

    else
        return nothing
    end
end

function safe_prev(str, i)
    i < firstindex(str) && return firstindex(str)
    return isvalid(str, i) ? i : prevind(str, i)
end

function safe_next(str, i)
    i > lastindex(str) && return lastindex(str)
    return isvalid(str, i) ? i : nextind(str, i)
end
