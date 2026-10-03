#=
Differential tests for `readtable` (#462).

`readtable(source, …)` builds its own `XLSXFile` and discards it. The oracle is
`gettable` on a file opened with `readxlsx`, which stays on the worksheet-cache path,
so any change to how `readtable` reads (or to `gettable` itself) must produce
exactly what the oracle produces: the same exception, or the same column labels,
column container types, row count, values (`isequal` and `typeof`), and the same
sequence of calls to `stop_in_row_function`. `eachtablerow` is checked against the
oracle's values too, so all three views of a table agree.

Corpora:
  1. a generated empty-row corpus: every sequence of 0–3 row states after a header,
     plus hand-made special sheets;
  2. every sheet of every workbook in `test/data`;
  3. randomised sparse sheets from fixed seeds;
  4. optionally, the bench fixtures (`XLSX_DIFF_FIXTURES=<dir>`).

Rows-deciding keywords (`header`, `stop_in_empty_row`, `keep_empty_rows`, `first_row`,
columns) are crossed in full; the others are drawn from a pairwise covering array.
Set `XLSX_FULL_DIFF=1` to cross every row-deciding combination with the whole
covering array (slow: tens of minutes).
=#

const _DIFF_FULL = get(ENV, "XLSX_FULL_DIFF", "0") == "1"
const _DIFF_CHECKS = Ref(0)        # comparisons run in the current section

# ── Minimal workbook builder ──────────────────────────────────────────────────
# Compact XML, as Excel writes it. Style 0 = General, 1 = date (numFmt 14),
# 2 = a fill only (for styled-but-empty cells).

const _DIFF_CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"""
const _DIFF_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""
const _DIFF_WB = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""
const _DIFF_WBRELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"""
const _DIFF_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs></styleSheet>"""

function _diff_build_xlsx(rows_xml::AbstractString, sst::Vector{String})::Vector{UInt8}
    sst_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" *
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"$(length(sst))\" uniqueCount=\"$(length(sst))\">" *
        join(isempty(s) ? "<si><t/></si>" : "<si><t xml:space=\"preserve\">$s</t></si>" for s in sst) * "</sst>"
    sheet_xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" *
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " *
        "xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">" *
        "<sheetData>$rows_xml</sheetData></worksheet>"
    io = IOBuffer()
    ZipArchives.ZipWriter(io) do w
        ZipArchives.zip_writefile(w, "[Content_Types].xml", codeunits(_DIFF_CT))
        ZipArchives.zip_writefile(w, "_rels/.rels", codeunits(_DIFF_RELS))
        ZipArchives.zip_writefile(w, "xl/workbook.xml", codeunits(_DIFF_WB))
        ZipArchives.zip_writefile(w, "xl/_rels/workbook.xml.rels", codeunits(_DIFF_WBRELS))
        ZipArchives.zip_writefile(w, "xl/styles.xml", codeunits(_DIFF_STYLES))
        ZipArchives.zip_writefile(w, "xl/sharedStrings.xml", codeunits(sst_xml))
        ZipArchives.zip_writefile(w, "xl/worksheets/sheet1.xml", codeunits(sheet_xml))
    end
    return take!(io)
end

# ── Empty-row corpus ──────────────────────────────────────────────────────────
# The table sits in B:D with its header on `header_row`. Column F is always empty
# (so a range widened to F reaches no data) and column G holds "outside" cells.
# Shared strings 0–2 are the labels a, b, c; 3 is ""; 4 is "NA"; 5+ are data strings.

const _DIFF_ROW_STATES = (:absent, :norow_cells, :styled_empty, :outside_only,
                          :empty_strings, :missing_strings, :formula_nocache, :data)

_diff_row(r, cells; attrs = "") = "<row r=\"$r\" spans=\"2:7\" x14ac:dyDescent=\"0.25\"$attrs>$cells</row>"

function _diff_state_row(state::Symbol, r::Int, sst::Vector{String})::String
    if state === :absent
        return ""
    elseif state === :norow_cells
        return "<row r=\"$r\"/>"
    elseif state === :styled_empty
        return _diff_row(r, "<c r=\"B$r\" s=\"2\"/><c r=\"C$r\" s=\"2\"/><c r=\"D$r\" s=\"2\"/>")
    elseif state === :outside_only
        return _diff_row(r, "<c r=\"G$r\"><v>$(100 + r)</v></c>")
    elseif state === :empty_strings
        return _diff_row(r, "<c r=\"B$r\" t=\"s\"><v>3</v></c>" *
                            "<c r=\"C$r\" t=\"inlineStr\"><is><t></t></is></c>" *
                            "<c r=\"D$r\" t=\"str\"><f>\"\"</f><v></v></c>")
    elseif state === :missing_strings
        return _diff_row(r, "<c r=\"B$r\" t=\"s\"><v>4</v></c><c r=\"C$r\" t=\"s\"><v>4</v></c>" *
                            "<c r=\"D$r\" t=\"inlineStr\"><is><t>NA</t></is></c>")
    elseif state === :formula_nocache
        return _diff_row(r, "<c r=\"B$r\"><f>1+$r</f></c><c r=\"C$r\" t=\"str\"><f>\"x\"&amp;$r</f></c>" *
                            "<c r=\"D$r\" s=\"1\"><f>TODAY()</f></c>")
    else # :data — Int-valued number, shared string, date; plus an outside cell
        push!(sst, "s$r")
        return _diff_row(r, "<c r=\"B$r\"><v>$r</v></c><c r=\"C$r\" t=\"s\"><v>$(length(sst) - 1)</v></c>" *
                            "<c r=\"D$r\" s=\"1\"><v>$(45000 + r)</v></c><c r=\"G$r\"><v>1</v></c>")
    end
end

_diff_base_sst() = ["a", "b", "c", "", "NA"]
_diff_header(r) = _diff_row(r, "<c r=\"B$r\" t=\"s\"><v>0</v></c><c r=\"C$r\" t=\"s\"><v>1</v></c><c r=\"D$r\" t=\"s\"><v>2</v></c>")

function _diff_corpus_sheet(states)::Vector{UInt8}
    sst = _diff_base_sst()
    xml = _diff_header(1) * join(_diff_state_row(s, i + 1, sst) for (i, s) in enumerate(states))
    return _diff_build_xlsx(xml, sst)
end

# Hand-made sheets for the cases a slot sequence can't express.
function _diff_special_sheets()::Vector{Pair{String,Vector{UInt8}}}
    out = Pair{String,Vector{UInt8}}[]
    d(r) = _diff_state_row(:data, r, sst)
    sst = _diff_base_sst()
    push!(out, "no rows" => _diff_build_xlsx("", copy(sst)))
    push!(out, "header only" => _diff_build_xlsx(_diff_header(1), copy(sst)))
    sst = _diff_base_sst()
    push!(out, "header absent, data from row 2" => _diff_build_xlsx(d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "header row empty" => _diff_build_xlsx(_diff_state_row(:styled_empty, 1, sst) * d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "header with gap" => _diff_build_xlsx(
        _diff_row(1, "<c r=\"B1\" t=\"s\"><v>0</v></c><c r=\"D1\" t=\"s\"><v>2</v></c>") * d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "title rows, header on 5" => _diff_build_xlsx(
        _diff_row(1, "<c r=\"A1\" t=\"inlineStr\"><is><t>Report</t></is></c>") *
        _diff_row(3, "<c r=\"A3\" t=\"inlineStr\"><is><t>notes</t></is></c>") *
        _diff_header(5) * d(6) * d(8) * d(9), sst))
    sst = _diff_base_sst()
    push!(out, "leading and trailing gaps" => _diff_build_xlsx(_diff_header(1) * d(5) * d(6) * d(12), sst))
    sst = _diff_base_sst()
    push!(out, "out-of-order rows" => _diff_build_xlsx(_diff_header(1) * d(4) * d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "duplicated row" => _diff_build_xlsx(_diff_header(1) * d(2) * d(2) * d(3), sst))
    sst = _diff_base_sst()
    push!(out, "mixed types per column" => _diff_build_xlsx(_diff_header(1) *
        _diff_row(2, "<c r=\"B2\"><v>1</v></c><c r=\"C2\" s=\"1\"><v>45000</v></c><c r=\"D2\" t=\"b\"><v>1</v></c>") *
        _diff_row(3, "<c r=\"B3\"><v>1.5</v></c><c r=\"C3\" s=\"1\"><v>45000.5</v></c><c r=\"D3\" t=\"e\"><v>#N/A</v></c>") *
        _diff_row(4, "<c r=\"B4\" t=\"inlineStr\"><is><t>x</t></is></c><c r=\"C4\" s=\"1\"><v>45001</v></c><c r=\"D4\" t=\"b\"><v>0</v></c>"), sst))
    return out
end

# ── Pairwise covering array for the secondary keywords ────────────────────────

# Greedy pairwise cover over `levels` (number of values per factor); deterministic.
function _diff_pairwise(levels::Vector{Int})::Vector{Vector{Int}}
    nf = length(levels)
    uncovered = Set{NTuple{4,Int}}()
    for i in 1:nf, j in (i+1):nf, a in 1:levels[i], b in 1:levels[j]
        push!(uncovered, (i, a, j, b))
    end
    candidates = vec(collect(Iterators.product((1:l for l in levels)...)))
    rows = Vector{Vector{Int}}()
    while !isempty(uncovered)
        best, best_n = candidates[1], -1
        for c in candidates
            n = 0
            for i in 1:nf, j in (i+1):nf
                (i, c[i], j, c[j]) in uncovered && (n += 1)
            end
            n > best_n && ((best, best_n) = (c, n))
        end
        for i in 1:nf, j in (i+1):nf
            delete!(uncovered, (i, best[i], j, best[j]))
        end
        push!(rows, collect(best))
    end
    return rows
end

# Secondary factors: enable_cache, infer_eltypes, normalizenames, column_labels,
# missing_strings, stop_in_row_function.
const _DIFF_SECONDARY_LEVELS = [2, 2, 2, 3, 4, 5]
const _DIFF_MS_OPTIONS = (nothing, "NA", ["NA", ""], "a")
const _DIFF_STOP_KINDS = (:none, :at2, :col1missing, :always, :never)

# ── Running one comparison ────────────────────────────────────────────────────

struct _DiffCall
    row::Int
    values::Vector{Any}
end

function _diff_stop_fn(kind::Symbol, log::Vector{_DiffCall})
    kind === :none && return nothing
    pred = kind === :at2 ? (r -> XLSX.row_number(r) >= 2) :
           kind === :col1missing ? (r -> ismissing(r[1])) :
           kind === :always ? (r -> true) : (r -> false)
    # Values by index: `TableRow` iterates but has no `length`, so it can't be collected.
    return r -> (push!(log, _DiffCall(XLSX.row_number(r),
                                      Any[r[i] for i in XLSX.table_column_numbers(r)])); pred(r))
end

# Outcome of one call: the DataTable, or the exception it threw.
_diff_run(f) = try f() catch e; e end

function _diff_describe(o)
    o isa Exception && return "exception $(typeof(o)): $(sprint(showerror, o))"
    return "table $(o.column_labels) $(map(typeof, o.data)) rows=$(isempty(o.data) ? 0 : length(o.data[1]))"
end

# First difference between two outcomes, or `nothing` when they agree exactly.
function _diff_compare(a, b)::Union{Nothing,String}
    if a isa Exception || b isa Exception
        (a isa Exception && b isa Exception) || return "one threw: $(_diff_describe(a)) vs $(_diff_describe(b))"
        typeof(a) == typeof(b) || return "exception types differ: $(typeof(a)) vs $(typeof(b))"
        sprint(showerror, a) == sprint(showerror, b) || return "exception messages differ: $(sprint(showerror, a)) vs $(sprint(showerror, b))"
        return nothing
    end
    a.column_labels == b.column_labels || return "labels differ: $(a.column_labels) vs $(b.column_labels)"
    length(a.data) == length(b.data) || return "column counts differ"
    for (ci, (ca, cb)) in enumerate(zip(a.data, b.data))
        typeof(ca) == typeof(cb) || return "column $ci type differs: $(typeof(ca)) vs $(typeof(cb))"
        length(ca) == length(cb) || return "column $ci length differs: $(length(ca)) vs $(length(cb))"
        for ri in eachindex(ca)
            x, y = ca[ri], cb[ri]
            (isequal(x, y) && typeof(x) == typeof(y)) ||
                return "column $ci row $ri differs: $(repr(x))::$(typeof(x)) vs $(repr(y))::$(typeof(y))"
        end
    end
    return nothing
end

function _diff_compare_calls(a::Vector{_DiffCall}, b::Vector{_DiffCall})::Union{Nothing,String}
    length(a) == length(b) || return "stop_in_row_function called $(length(a)) vs $(length(b)) times"
    for (i, (x, y)) in enumerate(zip(a, b))
        x.row == y.row || return "callback $i saw table row $(x.row) vs $(y.row)"
        (length(x.values) == length(y.values) &&
         all(isequal(p, q) && typeof(p) == typeof(q) for (p, q) in zip(x.values, y.values))) ||
            return "callback $i saw values $(x.values) vs $(y.values)"
    end
    return nothing
end

# `eachtablerow` gives untyped rows; its values must equal the oracle's (`isequal`).
function _diff_compare_rows(rows, oracle)::Union{Nothing,String}
    rows isa Exception && return oracle isa Exception ? nothing : "eachtablerow threw: $(_diff_describe(rows))"
    oracle isa Exception && return "eachtablerow succeeded where gettable threw"
    n = isempty(oracle.data) ? 0 : length(oracle.data[1])
    length(rows) == n || return "eachtablerow gave $(length(rows)) rows vs $n"
    for (ri, r) in enumerate(rows), ci in eachindex(oracle.data)
        isequal(r[ci], oracle.data[ci][ri]) || return "eachtablerow row $ri column $ci: $(repr(r[ci])) vs $(repr(oracle.data[ci][ri]))"
    end
    return nothing
end

# Known, pre-existing differences of `readtable(…; enable_cache=false)` (the streaming
# path) from the cache path, accepted only until Stage 5 of #462:
#  - a missing `first_row` throws "Row N not found in worksheet X." where the cache
#    path throws "Row N not found." (same exception type, same N);
#  - sheets with out-of-order or duplicated `<row r>` (which Excel never writes) are
#    read differently.
function _diff_nocache_allowed(oracle, subject, malformed::Bool)::Bool
    malformed && return true
    (oracle isa XLSX.XLSXError && subject isa XLSX.XLSXError) || return false
    a = match(r"^Row (\d+) not found\.$", oracle.msg)
    b = match(r"^Row (\d+) not found in worksheet .*\.$", subject.msg)
    return a !== nothing && b !== nothing && a[1] == b[1]
end

"""
Compare `readtable` with the `gettable` oracle for one keyword set. `xf` is the
oracle's open file; `bytes` the workbook. Returns a description of the first
mismatch, or `nothing`.
"""
function _diff_check(bytes::Vector{UInt8}, xf, sheet, cols, kw::NamedTuple, sec::Vector{Int};
                     check_rows::Bool = true, malformed::Bool = false)
    _DIFF_CHECKS[] += 1
    enable_cache = sec[1] == 1
    infer_eltypes = sec[2] == 1
    normalizenames = sec[3] == 2
    missing_strings = _DIFF_MS_OPTIONS[sec[5]]
    stop_kind = _DIFF_STOP_KINDS[sec[6]]

    # column_labels: nothing, the right length, or one too many (must throw identically)
    column_labels = nothing
    if sec[4] != 1
        probe = _diff_run(() -> isnothing(cols) ?
            XLSX.gettable(xf[sheet]; kw...) : XLSX.gettable(xf[sheet], cols; kw...))
        probe isa Exception && return nothing     # nothing more to learn for this case
        n = length(probe.column_labels) + (sec[4] == 3 ? 1 : 0)
        column_labels = [Symbol("L$i") for i in 1:n]
    end

    log_a, log_b = _DiffCall[], _DiffCall[]
    common(log) = (; kw..., column_labels, normalizenames, missing_strings,
                    stop_in_row_function = _diff_stop_fn(stop_kind, log))

    oracle = _diff_run(() -> isnothing(cols) ?
        XLSX.gettable(xf[sheet]; common(log_a)..., infer_eltypes) :
        XLSX.gettable(xf[sheet], cols; common(log_a)..., infer_eltypes))
    subject = _diff_run(() -> if isnothing(cols)
            sheet == 1 && sec[6] % 2 == 0 ?   # exercise readtable(source) too
                XLSX.readtable(IOBuffer(bytes); common(log_b)..., infer_eltypes, enable_cache) :
                XLSX.readtable(IOBuffer(bytes), sheet; common(log_b)..., infer_eltypes, enable_cache)
        else
            XLSX.readtable(IOBuffer(bytes), sheet, cols; common(log_b)..., infer_eltypes, enable_cache)
        end)

    # Known differences of the streaming path (enable_cache=false) from the cache
    # path, accepted until Stage 5 of #462 makes enable_cache a no-op for readtable.
    # Remove `_diff_nocache_allowed` then, and require exact agreement everywhere.
    !enable_cache && _diff_nocache_allowed(oracle, subject, malformed) && return nothing

    m = _diff_compare(oracle, subject)
    m === nothing || return "readtable: " * m
    m = _diff_compare_calls(log_a, log_b)
    m === nothing || return m

    if check_rows
        log_c = _DiffCall[]
        rows = _diff_run(() -> collect(isnothing(cols) ?
            XLSX.eachtablerow(xf[sheet]; common(log_c)...) :
            XLSX.eachtablerow(xf[sheet], cols; common(log_c)...)))
        m = _diff_compare_rows(rows, oracle)
        m === nothing || return m
    end
    return nothing
end

const _DIFF_COLUMNS = (nothing, "B:D", "B:C", "B:F", "C:C", "A:D", "H:I")

"""
Run the row-deciding keyword product on one workbook sheet. Secondary keywords come
from the covering array, rotated by `offset` so different sheets see different
pairings (all of them under `XLSX_FULL_DIFF=1`). Returns the mismatches found.
"""
function _diff_product(label, bytes, sheet, first_rows, columns, cover, offset;
                       malformed::Bool = false)::Vector{String}
    mismatches = String[]
    xf = _diff_run(() -> XLSX.readxlsx(IOBuffer(bytes)))
    if xf isa Exception
        # The oracle can't open it: readtable must fail the same way.
        e = _diff_run(() -> XLSX.readtable(IOBuffer(bytes), sheet))
        m = _diff_compare(xf, e)
        m === nothing || push!(mismatches, "$label: open: $m")
        return mismatches
    end
    k = offset
    for header in (true, false), stop_in_empty_row in (true, false), keep_empty_rows in (true, false),
        first_row in first_rows, cols in columns
        kw = (; header, stop_in_empty_row, keep_empty_rows, first_row)
        secs = _DIFF_FULL ? cover : [cover[mod1(k += 1, length(cover))]]
        for sec in secs
            m = _diff_check(bytes, xf, sheet, cols, kw, sec; malformed)
            m === nothing && continue
            push!(mismatches, "$label: cols=$(repr(cols)) $(kw) sec=$(sec): $m")
        end
    end
    return mismatches
end

# Prints the first mismatches; `XLSX_DIFF_LOG=<file>` appends all of them to a file.
function _diff_report(mismatches::Vector{String}, what::String, t0::Float64)
    println("Differential $what: $(_DIFF_CHECKS[]) checks, $(length(mismatches)) mismatches, ",
            round(time() - t0, digits = 1), " s")
    isempty(mismatches) && return
    log = get(ENV, "XLSX_DIFF_LOG", "")
    isempty(log) || open(io -> foreach(m -> println(io, what, " | ", m), mismatches), log, "a")
    println("Differential mismatches ($what), first $(min(20, length(mismatches))) of $(length(mismatches)):")
    foreach(m -> println("  ", m), first(mismatches, 20))
end

# ── Randomised sheets ─────────────────────────────────────────────────────────

function _diff_random_sheet(rng)::Vector{UInt8}
    sst = _diff_base_sst()
    header_row = rand(rng, 1:3)
    xml = IOBuffer()
    print(xml, _diff_header(header_row))
    for r in (header_row+1):(header_row+rand(rng, 0:12))
        state = rand(rng, (:absent, :absent, :norow_cells, :styled_empty, :outside_only,
                           :empty_strings, :missing_strings, :formula_nocache, :data, :data, :data, :sparse))
        if state === :sparse      # a random subset of B:D, mixed kinds
            cells = String[]
            for col in ("B", "C", "D")
                rand(rng) < 0.5 && continue
                push!(cells, rand(rng, (
                    "<c r=\"$col$r\"><v>$(rand(rng, -5:5))</v></c>",
                    "<c r=\"$col$r\"><v>$(rand(rng) * 10)</v></c>",
                    "<c r=\"$col$r\" s=\"1\"><v>$(45000 + rand(rng, 0:9))</v></c>",
                    "<c r=\"$col$r\" t=\"b\"><v>$(rand(rng, 0:1))</v></c>",
                    "<c r=\"$col$r\" t=\"inlineStr\"><is><t>i$r</t></is></c>",
                    "<c r=\"$col$r\" t=\"s\"><v>$(rand(rng, 3:4))</v></c>",
                    "<c r=\"$col$r\" t=\"e\"><v>#DIV/0!</v></c>",
                    "<c r=\"$col$r\" s=\"2\"/>")))
            end
            print(xml, isempty(cells) ? "" : _diff_row(r, join(cells)))
        else
            print(xml, _diff_state_row(state, r, sst))
        end
    end
    return _diff_build_xlsx(String(take!(xml)), sst)
end

# ── Tests ─────────────────────────────────────────────────────────────────────

@testset "Differential: readtable vs gettable" begin
    cover = _diff_pairwise(_DIFF_SECONDARY_LEVELS)

    @testset "pairwise cover is complete" begin
        levels = _DIFF_SECONDARY_LEVELS
        for i in eachindex(levels), j in (i+1):length(levels), a in 1:levels[i], b in 1:levels[j]
            @test any(r -> r[i] == a && r[j] == b, cover)
        end
    end

    @testset "empty-row corpus" begin
        _DIFF_CHECKS[] = 0; t0 = time()
        mismatches = String[]
        n = 0
        for len in 0:3, states in Iterators.product(ntuple(_ -> _DIFF_ROW_STATES, len)...)
            n += 1
            bytes = _diff_corpus_sheet(collect(states))
            append!(mismatches, _diff_product("corpus $(collect(states))", bytes, 1,
                                              (nothing, 1, 2, 3, 4, 5), _DIFF_COLUMNS, cover, n))
        end
        @test n == 585
        for (name, bytes) in _diff_special_sheets()
            n += 1
            append!(mismatches, _diff_product("special \"$name\"", bytes, 1,
                                              (nothing, 1, 2, 3, 5, 6, 7, 13), _DIFF_COLUMNS, cover, n;
                                              malformed = name in ("out-of-order rows", "duplicated row")))
        end
        _diff_report(mismatches, "empty-row corpus", t0)
        @test isempty(mismatches)
    end

    @testset "randomised sheets" begin
        _DIFF_CHECKS[] = 0; t0 = time()
        mismatches = String[]
        for seed in 1:500
            rng = Random.MersenneTwister(seed)
            bytes = _diff_random_sheet(rng)
            for _ in 1:8
                cols = rand(rng, _DIFF_COLUMNS)
                kw = (; header = rand(rng, Bool), stop_in_empty_row = rand(rng, Bool),
                        keep_empty_rows = rand(rng, Bool), first_row = rand(rng, (nothing, 1:8...)))
                sec = rand(rng, cover)
                xf = XLSX.readxlsx(IOBuffer(bytes))
                m = _diff_check(bytes, xf, 1, cols, kw, sec)
                m === nothing || push!(mismatches, "seed $seed cols=$(repr(cols)) $kw sec=$sec: $m")
            end
        end
        _diff_report(mismatches, "randomised", t0)
        @test isempty(mismatches)
    end

    @testset "test/data workbooks" begin
        _DIFF_CHECKS[] = 0; t0 = time()
        mismatches = String[]
        n = 0
        for file in sort(readdir(data_directory))
            any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
            bytes = read(joinpath(data_directory, file))
            xf = _diff_run(() -> XLSX.readxlsx(IOBuffer(bytes)))
            sheets = xf isa Exception ? [1] : collect(1:XLSX.sheetcount(xf))
            # Large workbooks get the row-deciding booleans only, unless XLSX_FULL_DIFF=1.
            _diff_run(() -> XLSX.readtable(IOBuffer(bytes), 1))
            slow = !_DIFF_FULL && (@elapsed _diff_run(() -> XLSX.readtable(IOBuffer(bytes), 1))) > 0.02
            for sheet in sheets
                n += 1
                dim = xf isa Exception ? nothing : _diff_run(() -> XLSX.get_dimension(xf[sheet]))
                columns = Any[nothing]
                first_rows = Any[nothing, 1, 2]
                if slow
                    append!(mismatches, _diff_product("$file sheet $sheet", bytes, sheet,
                                                      Any[nothing], columns, cover, n))
                    continue
                end
                if dim isa XLSX.CellRange
                    c1, c2 = XLSX.column_number(dim.start), XLSX.column_number(dim.stop)
                    r2 = XLSX.row_number(dim.stop)
                    col(i) = XLSX.encode_column_number(i)
                    append!(columns, ["$(col(c1)):$(col(c2))", "$(col(c1)):$(col(c1))",
                                      "$(col(c1)):$(col(c2 + 2))"])
                    append!(first_rows, [r2, r2 + 1])
                end
                append!(mismatches, _diff_product("$file sheet $sheet", bytes, sheet,
                                                  first_rows, columns, cover, n))
            end
        end
        _diff_report(mismatches, "test/data", t0)
        @test isempty(mismatches)
    end

    fixtures = get(ENV, "XLSX_DIFF_FIXTURES", "")
    if !isempty(fixtures)
        @testset "bench fixtures ($fixtures)" begin
            _DIFF_CHECKS[] = 0; t0 = time()
            mismatches = String[]
            for file in sort(readdir(fixtures))
                startswith(file, "xl_") || continue
                bytes = read(joinpath(fixtures, file))
                xf = XLSX.readxlsx(IOBuffer(bytes))
                for (i, sec) in enumerate(cover)
                    kw = (; header = true, stop_in_empty_row = isodd(i), keep_empty_rows = i % 4 < 2, first_row = 5)
                    m = _diff_check(bytes, xf, 2, "A:CF", kw, sec; check_rows = false)
                    m === nothing || push!(mismatches, "$file $kw sec=$sec: $m")
                end
            end
            _diff_report(mismatches, "bench fixtures", t0)
            @test isempty(mismatches)
        end
    end
end
