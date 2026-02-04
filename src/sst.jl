
SharedStringTable() = SharedStringTable(Vector{String}(), Dict{String, Int64}(), false)

@inline get_sst(wb::Workbook) = wb.sst
@inline get_sst(xl::XLSXFile) = get_sst(get_workbook(xl))
@inline Base.length(sst::SharedStringTable) = length(sst.shared_strings)
@inline Base.isempty(sst::SharedStringTable) = isempty(sst.shared_strings)

# Checks if string is inside shared string table.
# Returns `nothing` if it's not in the shared string table.
# Returns the index of the string in the shared string table. The index is 0-based.
function get_shared_string_index(sst::SharedStringTable, str_formatted::String)# :: Union{Nothing, Int}
    !sst.is_loaded && throw(XLSXError("Can't query shared string table because it's not loaded into memory."))

    #using a Dict is much more efficient than the findfirst approach especially on large datasets
    k = str_formatted
    if haskey(sst.index, k)
        return sst.index[k]
    else
        return nothing
    end

end
function create_new_sst(wb::Workbook, sst::SharedStringTable)
    if !sst.is_loaded
        sst.is_loaded = true

        # add relationship
        #<Relationship Id="rId16" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
        add_relationship!(wb, "sharedStrings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")

        # add Content Type <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>
        ctype_root = xmlroot(get_xlsxfile(wb), "[Content_Types].xml")[end]
        XML.tag(ctype_root) != "Types" && throw(XLSXError("Something wrong here!"))
        override_node = XML.Element("Override";
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            PartName = "/xl/sharedStrings.xml"
        )
        push!(ctype_root, override_node)
        init_sst_index(sst)
    end
end
function add_to_sst!(ss::SharedStringTable, si_xml::String)::Int64
    
    # Check for match
    ind = get(ss.index, si_xml, nothing)
    if ind !== nothing
        return ind  # Found exact match
    end
    
    # No match found, add new entry
    push!(ss.shared_strings, si_xml)
    new_idx = length(ss.shared_strings)-1  # 0-based index

    ss.index[si_xml] = new_idx

    if new_idx ∉ get_shared_string_index(ss, si_xml)
        throw(XLSXError("Inconsistent state after adding a string to the Shared String Table."))
    end

    return new_idx
end

function add_formatted_string!(sst::SharedStringTable, str_formatted::String; mylock::Union{Nothing,ReentrantLock}=nothing) :: Int64
    ind = get_shared_string_index(sst, str_formatted)
    local new_index::Int
    if ind !== nothing
        # it's already in the table
        return ind  # Found exact match
    end
    if isnothing(mylock)
        new_index = add_to_sst!(sst, str_formatted)
    else
        lock(mylock) do
            new_index = add_to_sst!(sst, str_formatted)
        end
    end
    return new_index
end

# Adds a string to shared string table. Returns the 0-based index of the shared string in the shared string table.
function add_formatted_string!(wb::Workbook, str_formatted::String; mylock::Union{Nothing,ReentrantLock}=nothing) :: Int64
#    !is_writable(get_xlsxfile(wb)) && throw(XLSXError("XLSXFile instance is not writable."))
    if isempty(str_formatted)
        throw(XLSXError("Can't add empty string to Shared String Table."))
    end
    sst = get_sst(wb)

        # if got to this point, the file was opened as template but doesn't have a Shared String Table.
        # Will create a new one.
    if !sst.is_loaded
        if isnothing(mylock)
            create_new_sst(wb, sst)
        else
            lock(mylock) do # ensure thread-safety if multiple threads are trying to add inlineStrings
                create_new_sst(wb, sst)
            end
        end
    end
    
    return add_formatted_string!(sst, str_formatted; mylock)
end

# allow to write cells containing only whitespace characters or with leading or trailing whitespace.
function add_shared_string!(wb::Workbook, str_unformatted::AbstractString; mylock::Union{Nothing,ReentrantLock}=nothing) :: Int
    if startswith(str_unformatted, ' ') || endswith(str_unformatted, ' ') || contains(str_unformatted, '\n')  || contains(str_unformatted, "  ")
        str_formatted = string("<si>\n  <t xml:space=\"preserve\">", XML.escape(str_unformatted), "</t>\n</si>")
    else
        str_formatted = string("<si>\n  <t>", XML.escape(str_unformatted), "</t>\n</si>")
    end
    return add_formatted_string!(wb, str_formatted; mylock)
end

function sst_load!(workbook::Workbook)
    chunksize=1000
    sst = get_sst(workbook)
    if !sst.is_loaded

        relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        if has_relationship_by_type(workbook, relationship_type)
            sst_chan = stream_ssts(open_internal_file_stream(get_xlsxfile(workbook), "xl/sharedStrings.xml")[end], chunksize)
            load_sst_table!(workbook, sst_chan, chunksize, Threads.nthreads())
            init_sst_index(sst)
            
            return
        end

        throw(XLSXError("Shared Strings Table not found for this workbook."))
    end
end
@inline _is_tag(n::String, tag::String) = n == tag
@inline _is_tag(n::Nothing, tag::String) = false
function produce_sstchunks(out, n, ssts, chunksize)
    i=0
    while !isnothing(n)
        if _is_tag(n.tag, "si")
            i += 1
            ssts[i] = SstToken(n, i)
        end
        if i >= chunksize
            put!(out, copy(ssts))
            i=0
        end
        n = XML.next(n)
    end
    if i>0 # handle last incomplete chunk
        put!(out, copy(ssts[1:i]))
    end
    return out
end
function stream_ssts(n::XML.LazyNode, chunksize::Int; channel_size::Int=1 << 8)
    n = XML.next(n)
    ssts = Vector{SstToken}(undef, chunksize)
    Channel{Vector{SstToken}}(channel_size) do out
        produce_sstchunks(out, n, ssts, chunksize)
    end
end

function process_sst(sst::SstToken)
    el = sst.n
    i = sst.idx

    if XML.nodetype(el) != XML.Text
        XML.tag(el) != "si" && throw(XLSXError("Unsupported node $(XML.tag(el)) in sst table."))
        sst = Sst(XML.write(el), i)
        return sst

    end

end

function load_sst_table!(wb::Workbook, chan::Channel, chunksize::Int, nthreads::Int)
    sst_table = get_sst(wb)
    sst_table.is_loaded=true

    sst_results = Channel{Vector{Sst}}(1 << 8)
    all_ssts = Vector{Tuple{Int,Sst}}()

    consumer = @async begin
        for ssts in sst_results        
            for sst in ssts
                push!(all_ssts, (sst.idx, sst))
            end
        end    

        sort!(all_ssts, by = x -> x[1])
    
        empty!(sst_table.index)
        for sst in all_ssts
            add_formatted_string!(sst_table, sst[end].formatted)
        end
    
    end

    # Producer tasks
    @sync for _ in 1:nthreads
        Threads.@spawn begin
            for ssts in chan
                chunk = Vector{Sst}(undef, chunksize)
                sst_count=0
                for tok in ssts
                    sst_count += 1
                    chunk[sst_count] = process_sst(tok)
                    if sst_count == chunksize
                        put!(sst_results, copy(chunk))
                        sst_count=0
                    end
                end
                if sst_count>0 # handle last incomplete chunk
                    put!(sst_results, copy(chunk[1:sst_count]))
                end
            end
        end
    end
    
    close(sst_results)

    wait(consumer)  # ensure consumer is done

#   sst_table.is_loaded=true

end

# Checks whether this workbook has a Shared String Table.
function has_sst(workbook::Workbook) :: Bool
    relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    return has_relationship_by_type(workbook, relationship_type)
end

# Helper function to gather unformatted text from Excel data files.
# It looks at all children of `el` for tag name `t` and returns
# a join of all the strings found.
function unformatted_text(el::XML.LazyNode) :: String
    io = IOBuffer()
    gather_strings!(io, el)
    return XML.unescape(String(take!(io)))
end

function gather_strings!(io::IOBuffer, e::XML.LazyNode)
    tag = XML.tag(e)
    
    # Skip phonetic hints entirely
    tag == "rPh" && return nothing
    
    if tag == "t"
        children = XML.children(e)
        n = length(children)
        
        if n == 1
            c = children[1]
            write(io, XML.is_simple(c) ? XML.simple_value(c) : XML.value(c))
        elseif n == 0
            val = XML.value(e)
            !isnothing(val) && write(io, XML.is_simple(e) ? XML.simple_value(e) : val)
        else
            throw(XLSXError("Unexpected number of children in <t>: $n. Expected 0 or 1."))
        end
    else
        # Recurse into children for all other tags
        for ch in XML.children(e)
            gather_strings!(io, ch)
        end
    end
    
    return nothing
end

# Looks for a string inside the Shared Strings Table (sst).
# `index` starts at 0.
@inline function sst_unformatted_string(wb::Workbook, index::Int)::String
    sst_load!(wb)
    uss = get_sst(wb).shared_strings[index+1]
    return unformatted_text(parse(XML.LazyNode, uss))
end

@inline sst_unformatted_string(xl::XLSXFile, index::Int) :: String = sst_unformatted_string(get_workbook(xl), index)
@inline sst_unformatted_string(ws::Worksheet, index::Int) :: String = sst_unformatted_string(get_xlsxfile(ws), index)
@inline sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) :: String = sst_unformatted_string(target, parse(Int, index_str))

# init the index table
function init_sst_index(sst::SharedStringTable)
    empty!(sst.index)
    for i in 1:length(sst.shared_strings)
        ind = get(sst.index, sst.shared_strings[i], nothing)
        if ind === nothing
            sst.index[sst.shared_strings[i]] = i-1
        else
            sst.index[sst.shared_strings[i]] = ind-1
        end
    end
end

#=
This is the required order of attributes in the xml:
- <rFont> — Font name
- <charset> — Character set # Not required/used here
- <family> — Font family # Not required/used here
- <b> — Bold
- <i> — Italic
- <strike> — Strikethrough
- <outline>
- <shadow>
- <condense>
- <extend>
- <color> — Color
- <sz> — Font size
- <u> — Underline
- <vertAlign> — Superscript/subscript
- <scheme> — Font scheme (major/minor) # Not required/used here
=#

 """
    richTextRunToXML(run::RichTextRun) -> String

Convert an RichTextRun to XML format for Excel shared strings.
Each rich text shared string may have multiple runs to allow 
heterogeneous formatting within a single cell.
"""
function richTextRunToXML!(io::IO, run::RichTextRun)
    write(io, "<r>")

    atts=run.atts

    if !isnothing(atts)

        # Build rPr (run properties) if any formatting is specified
        props = IOBuffer()
    
        if haskey(atts, :name)
            write(props, "<rFont val=\"$(atts[:name])\"/>")
        end
    
        if haskey(atts, :bold) && atts[:bold] == true
            write(props, "<b/>")
        end
    
        if haskey(atts, :italic) && atts[:italic] == true
            write(props, "<i/>")
        end
    
        if haskey(atts, :strike) && atts[:strike] == true
            write(props, "<strike/>")
        end
    
        if haskey(atts, :color)
            write(props, "<color rgb=\"$(get_color(atts[:color]))\"/>")
        end
    
        if haskey(atts, :size)
            write(props, "<sz val=\"$(atts[:size])\"/>")
        end
    
        if haskey(atts, :under) && atts[:under] == true
            write(props, "<u/>")
        end
    
        if haskey(atts, :vertAlign)
            write(props, "<vertAlign val=\"$(atts[:vertAlign])\"/>")
        end
    
    # emit rPr if needed
        if position(props) > 0
            write(io, "<rPr>", String(take!(props)), "</rPr>")
        end

end

    # whitespace rules
    needs_preserve =
        startswith(run.text, " ") ||
        endswith(run.text, " ") ||
        contains(run.text, '\n') ||
        contains(run.text, "  ")

    escaped = XML.escape(run.text)

    if needs_preserve
        write(io, "<t xml:space=\"preserve\">", escaped, "</t>")
    else
        write(io, "<t>", escaped, "</t>")
    end

    write(io, "</r>")

    return nothing
#    return String(take!(io))
end

function richTextStringtoXML(rts::RichTextString)
    xml = IOBuffer()
    write(xml, "<si>")
    for r in rts.runs
        richTextRunToXML!(xml, r)
    end
    write(xml, "</si>")
    return String(take!(xml))
end
function richTextString(runs::Vector{RichTextRun})
    isempty(runs) && throw(XLSXError("Cannot create a RichTextString with no RichTextRuns"))
    t = join([x.text for x in runs])
    return RichTextString(t, runs)
end

function ssToRuns(args...)

    error("""
    The use of styled strings requires the StyledStrings.jl package.
    
    Please install and load it with:
        using Pkg
        Pkg.add("StyledStrings")
        using StyledStrings
    
    Then retry your XLSX call with your styled string.

    Alternatively, use the XLSX type RichTextString directly,
    """)
end