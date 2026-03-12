module StyledStringsSstsExt 

using XLSX

# StyledStrings is only available in Julia 1.11+
@static if VERSION >= v"1.11-"

    using StyledStrings

    # Import from StyledStrings
    import StyledStrings: load_customisations!, getface, Face, FACES, SimpleColor, HTML_BASIC_COLORS

    # Import from XLSX
    import XLSX: setdata!, RichTextString, RichTextRun

    # Colors.jl is already a dependency of XLSX
    #import Colors: RGB


    setdata!(sheet::Worksheet, ref::CellRef, ss::AnnotatedString{T}) where T = setdata!(sheet, ref, RichTextString(_ssToRuns(ss)))


    """
        _ssToRuns(s::Union{<:AnnotatedString{>:Face}, SubString{<:AnnotatedString{>:Face}}}) -> Vector{XLSX.RichTextRun}
        _ssToRuns(s::Union{<:AnnotatedString, SubString{<:AnnotatedString}})                 -> Vector{XLSX.RichTextRun}

    Converts a `StyledString` to a vector of `RichTextRun`s.
    """
    function _ssToRuns end

    @static if VERSION >= v"1.14-"
        const _SS_ARG = Union{<:AnnotatedString{>:Face}, SubString{<:AnnotatedString{>:Face}}}
    else
        const _SS_ARG = Union{<:AnnotatedString, SubString{<:AnnotatedString}}
    end

    function _ssToRuns(s::_SS_ARG)

        runs = RichTextRun[]

        StyledStrings.load_customisations!()

        for (str, styles) in Base.eachregion(s)
            face = StyledStrings.getface(styles)
            d =_ss_style(face)
            push!(runs, RichTextRun(String(str), collect(d)))
        end

        return runs
    end

    function _ss_color(color::StyledStrings.SimpleColor)

        if color.value isa Symbol
            if color.value === :default
                return ""
            elseif (fg = get(StyledStrings.FACES.current[], color.value, StyledStrings.getface()).foreground) != StyledStrings.SimpleColor(color.value)
                return _ss_color(fg)
            else
                return _ss_color(get(StyledStrings.HTML_BASIC_COLORS, color.value, StyledStrings.SimpleColor(:default)))
            end

        else
            r, g, b = color.value.r, color.value.g, color.value.b
            io=IOBuffer()
            print(io, '#')
            r < 0x10 && print(io, '0')
            print(io, string(r, base=16))
            g < 0x10 && print(io, '0')
            print(io, string(g, base=16))
            b < 0x10 && print(io, '0')
            print(io, string(b, base=16))
            return String(take!(io))
        end
    end

    """
        _ss_style(face::StyledStrings.Face) -> Dict{Symbol, Any}

    Creates a dictionary of Excel font attributes from a StyledString `face`.

    Returns a Dict of (attribute => value).
    """
    function _ss_style(face::StyledStrings.Face)
        d = Dict{Symbol, Any}()

        if !isnothing(face.font)
            # monospace is default so allow Excel's default here, too.
            if face.font != "monospace"
                # Otherwise just pass through font name verbatim and without validation.
                d[:name] = face.font
            end
        end

        if !isnothing(face.height) # held as "deci-pt"
            d[:size] = face.height ÷ 10
        end

        if !isnothing(face.weight)
            if face.weight in [:medium, :semibold, :bold, :extrabold, :black]
                d[:bold] = true
            end
        end

        if !isnothing(face.slant)
            if face.slant in [:italic, :oblique]
                d[:italic] = true
            end
        end

        if !isnothing(face.foreground)
            c = _ss_color(face.foreground)
            if c != ""
                d[:color] = _ss_color(face.foreground)
            end
        end

        # Ignore background and invert.

        if !isnothing(face.underline)
            # Ignore underline color

            if face.underline isa Tuple # Color and style
                d[:under] = true

            elseif face.underline isa StyledStrings.SimpleColor
                d[:under] = true

            else # must be a Bool
                if face.underline
                    d[:under] = true
                end

            end
        end

        if !isnothing(face.strikethrough)
            if face.strikethrough
                d[:strike] = true
            end
        end

        return d
    end
    
end # @static if VERSION >= v"1.11-"

end # module
