
SharedStringTable() = SharedStringTable(Vector{String}(), Dict{UInt64, Vector{Int64}}(), false)

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
    k=hash(str_formatted)
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
function add_to_sst!(ss::SharedStringTable, si_xml::String)::Int
    xml_hash = hash(si_xml)
    
    # Check all indices with same hash
    indices = get(ss.index, xml_hash, nothing)
    if indices !== nothing
        for idx in indices
            if ss.shared_strings[idx+1] == si_xml
                return idx  # Found exact match
            end
        end
    end
    
    # No match found, add new entry
    push!(ss.shared_strings, si_xml)
    new_idx = length(ss.shared_strings)-1  # 0-based index

    if indices === nothing
        ss.index[xml_hash] = [new_idx]
    else
        push!(indices, new_idx)
    end

    if new_idx ∉ get_shared_string_index(ss, si_xml)
        throw(XLSXError("Inconsistent state after adding a string to the Shared String Table."))
    end

    return new_idx
end

function add_formatted_string!(sst::SharedStringTable, str_formatted::String; mylock::Union{Nothing,ReentrantLock}=nothing) :: Int
    indices = get_shared_string_index(sst, str_formatted)
    local new_index::Int
    if indices !== nothing
        # it's already in the table
        for idx in indices
            if sst.shared_strings[idx+1] == str_formatted
                return idx  # Found exact match
            end
        end
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
function add_formatted_string!(wb::Workbook, str_formatted::String; mylock::Union{Nothing,ReentrantLock}=nothing) :: Int
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
function add_shared_string!(wb::Workbook, str_unformatted::AbstractString) :: Int
    if startswith(str_unformatted, ' ') || endswith(str_unformatted, ' ') || contains(str_unformatted, '\n')
        str_formatted = string("<si><t xml:space=\"preserve\">", XML.escape(str_unformatted), "</t></si>")
    else
        str_formatted = string("<si><t>", XML.escape(str_unformatted), "</t></si>")
    end
    return add_formatted_string!(wb, str_formatted)
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
function stream_ssts(n::XML.LazyNode, chunksize::Int; channel_size::Int=1 << 10)
    n = XML.next(n)
    ssts = Vector{SstToken}(undef, chunksize)
    i=0
    idx=0
    Channel{Vector{SstToken}}(channel_size) do out
        while !isnothing(n)
            if n.tag == "si"
                i += 1
                idx += 1
                ssts[i] = SstToken(n, idx)
            end
            if i >= chunksize
                put!(out, copy(ssts))
                i=0
            end
            n = XML.next(n)
        end
        if i>0 # handle last incomplete chunk
            put!(out, ssts[1:i])
        end
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

    sst_results = Channel{Vector{Sst}}(1 << 10)
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
                    put!(sst_results, chunk[1:sst_count])
                end
            end
        end
    end
    
    close(sst_results)

    wait(consumer)  # ensure consumer is done
   
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

    function gather_strings!(v::IOBuffer, e::XML.LazyNode)
        tag = XML.tag(e)
        children = XML.children(e)

        if tag == "t"
            n = length(children)

            if n == 1
                c = children[1]
                write(v, XML.is_simple(c) ? XML.simple_value(c) : XML.value(c))

            elseif n == 0
                val = XML.value(e)
                if !isnothing(val)
                    write(v, XML.is_simple(e) ? XML.simple_value(e) : val)
                end

            else
                throw(XLSXError("Unexpected number of children in <t>: $n. Expected 0 or 1."))
            end
        #end

        # Skip recursion early
        elseif tag != "rPh"
            for ch in children
                gather_strings!(v, ch)
            end
        end

        return nothing
    end
    #=
    function gather_strings!(v::Vector{String}, e::XML.LazyNode)
        if XML.tag(e) == "t"
            c=XML.children(e)
            if length(c) == 1
                push!(v, XML.is_simple(c[1]) ? XML.simple_value(c[1]) : XML.value(c[1]))
            elseif length(c) == 0
                push!(v, isnothing(XML.value(e)) ? "" : XML.is_simple(e) ? XML.simple_value(e) : XML.value(e))
            else
                throw(XLSXError("Unexpected number of children in <t> node: $(length(c)). Expected 0 or 1."))
            end
        end

        if XML.tag(e) != "rPh"
            for ch in XML.children(e)
                # recursively gather strings from children
                gather_strings!(v, ch)
            end 
        end

        nothing
    end
    =#
    v_string = IOBuffer()
    gather_strings!(v_string, el)

    return XML.unescape(String(take!(v_string)))
end

# Looks for a string inside the Shared Strings Table (sst).
# `index` starts at 0.
@inline function sst_unformatted_string(wb::Workbook, index::Int)::String
    sst_load!(wb)
    uss = get_sst(wb).shared_strings[index+1]
    return unformatted_text(parse(XML.LazyNode, uss))
end

# Looks for a formatted string inside the Shared Strings Table (sst).
# `index` starts at 0.
@inline function sst_formatted_string(wb::Workbook, index::Int)
    sst_load!(wb)
    return get_sst(wb).shared_strings[index+1]
end

@inline sst_unformatted_string(xl::XLSXFile, index::Int) :: String = sst_unformatted_string(get_workbook(xl), index)
@inline sst_unformatted_string(ws::Worksheet, index::Int) :: String = sst_unformatted_string(get_xlsxfile(ws), index)
@inline sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) :: String = sst_unformatted_string(target, parse(Int, index_str))
#function sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) :: String
#    return sst_unformatted_string(target, parse(Int, index_str))
#end


# init the index table
function init_sst_index(sst::SharedStringTable)
    empty!(sst.index)
    for i in 1:length(sst.shared_strings)
        xmlhash = hash(sst.shared_strings[i])
        indices = get(sst.index, xmlhash, nothing)
        if indices === nothing
            sst.index[xmlhash] = [i]
        else
            push!(indices, i)
        end
    end
end
