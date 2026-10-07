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

# With `readtable` on the cache path (`XLSX._READTABLE_VALUE_ROWS[] = false`),
# `enable_cache=false` streams the sheet, which reads sheets with out-of-order or
# duplicated `<row r>` (never written by Excel) differently from the cache. That is
# the one accepted difference. With value rows (the default), `enable_cache` has no
# effect and such sheets fall back to the cache, so agreement must be exact.
function _diff_nocache_allowed(oracle, subject, malformed::Bool)::Bool
    return malformed && !XLSX._READTABLE_VALUE_ROWS[]
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

    # The one known difference of the streaming path (see `_diff_nocache_allowed`)
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

    # Every check runs with readtable on its value rows (the default, #462) and on the
    # worksheet cache, so neither path can drift from `gettable`.
    for value_rows in (true, false)
      @testset "readtable via $(value_rows ? "value rows" : "the cache")" begin
        XLSX._READTABLE_VALUE_ROWS[] = value_rows
        try
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
        finally
            XLSX._READTABLE_VALUE_ROWS[] = true
        end
      end
    end
end

# ── Value rows vs cache rows, row by row (#462) ───────────────────────────────

# Rows of `bytes`' sheet `sheet` from a `ValueRowIterator` (decoding `cols`, or all
# columns) and from the worksheet cache, or the exceptions they threw.
function _both_row_sources(bytes, sheet, cols)
    value_rows = _diff_run() do
        xf = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, true, false; split_sheets = false)
        ws = xf[sheet]
        raw = xf.data[XLSX.get_relationship_target_by_id("xl", XLSX.get_workbook(xf), ws.relationship_id)]
        collect(XLSX.ValueRowIterator(ws, raw, cols))
    end
    cache_rows = _diff_run(() -> collect(XLSX.eachrow(XLSX.readxlsx(IOBuffer(bytes))[sheet])))
    return value_rows, cache_rows
end

# First difference between value rows and cache rows, or `nothing`.
function _compare_row_sources(value_rows, cache_rows, cols)
    if value_rows isa Exception || cache_rows isa Exception
        value_rows isa XLSX._RowsNotAscending && return nothing   # readtable falls back to the cache
        return _diff_compare(cache_rows, value_rows)
    end
    length(value_rows) == length(cache_rows) || return "$(length(value_rows)) value rows vs $(length(cache_rows)) cache rows"
    for (v, c) in zip(value_rows, cache_rows)
        XLSX.row_number(v) == XLSX.row_number(c) || return "row numbers $(XLSX.row_number(v)) vs $(XLSX.row_number(c))"
        r = XLSX.row_number(c)
        isempty(v) == isempty(c) || return "row $r: isempty $(isempty(v)) vs $(isempty(c))"
        ws = XLSX.get_worksheet(c)
        check = if isnothing(cols)
            present = XLSX._present_columns(c)
            XLSX._present_columns(v) == present || return "row $r: present columns $(XLSX._present_columns(v)) vs $present"
            vcat(present, isempty(present) ? [1] : [maximum(present) + 1])   # plus an absent column
        else
            collect(cols)
        end
        for col in check
            a, b = XLSX._row_value(ws, v, col), XLSX._row_value(ws, c, col)
            (isequal(a, b) && typeof(a) == typeof(b)) || return "row $r col $col: $(repr(a)) vs $(repr(b))"
            la, lb = String[], String[]
            XLSX._push_header_label!(la, ws, v, col)
            XLSX._push_header_label!(lb, ws, c, col)
            la == lb || return "row $r col $col: header label $la vs $lb"
        end
    end
    return nothing
end

@testset "value rows match cache rows" begin
    sheets = Tuple{String,Vector{UInt8},Int}[]
    for len in 0:3, states in Iterators.product(ntuple(_ -> _DIFF_ROW_STATES, len)...)
        push!(sheets, ("corpus $(collect(states))", _diff_corpus_sheet(collect(states)), 1))
    end
    append!(sheets, [("special \"$k\"", v, 1) for (k, v) in _diff_special_sheets()])
    append!(sheets, [("random seed $s", _diff_random_sheet(Random.MersenneTwister(s)), 1) for s in 1:200])
    for file in sort(readdir(data_directory))
        any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
        bytes = read(joinpath(data_directory, file))
        xf = try XLSX.readxlsx(IOBuffer(bytes)) catch; continue end
        wb = XLSX.get_workbook(xf)
        for (i, ws) in enumerate(wb.sheets)
            XLSX.is_chartsheet(wb, ws.name) || push!(sheets, ("$file sheet $i", bytes, i))
        end
    end
    fixtures = get(ENV, "XLSX_DIFF_FIXTURES", "")
    if !isempty(fixtures)
        for file in sort(readdir(fixtures))
            startswith(file, "xl_") && push!(sheets, ("$file", read(joinpath(fixtures, file)), 2))
        end
    end

    mismatches = String[]
    for (label, bytes, sheet) in sheets, cols in (nothing, 2:4, 1:7, 3:3)
        m = _compare_row_sources(_both_row_sources(bytes, sheet, cols)..., cols)
        m === nothing || push!(mismatches, "$label cols=$cols: $m")
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test length(sheets) > 900

    # the malformed sheets make value rows hand over to the cache
    for (name, bytes) in _diff_special_sheets()
        name in ("out-of-order rows", "duplicated row") || continue
        @test _both_row_sources(bytes, 1, nothing)[1] isa XLSX._RowsNotAscending
    end
end

@testset "value rows: invalid cells outside the read area" begin
    # By design (src/valuerows.jl), readtable decodes only the cells it returns, so an
    # invalid cell elsewhere raises no error; the cache path decodes every cell.
    header = "<row r=\"1\"><c r=\"B1\" t=\"inlineStr\"><is><t>a</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>b</t></is></c></row>"
    good(r) = "<row r=\"$r\"><c r=\"B$r\"><v>$r</v></c><c r=\"C$r\"><v>$(10r)</v></c></row>"
    bad_bool(ref) = "<c r=\"$ref\" t=\"b\"><v>2</v></c>"
    outside_column = _diff_build_xlsx(header * good(2) * "<row r=\"3\"><c r=\"B3\"><v>3</v></c><c r=\"C3\"><v>30</v></c>$(bad_bool("E3"))</row>", String[])
    # the table ends at the gap before row 6; row 6 is read to find that end, row 7 never is
    after_table = _diff_build_xlsx(header * good(2) * good(3) * good(6) * "<row r=\"7\">$(bad_bool("B7"))</row>", String[])
    inside = _diff_build_xlsx(header * good(2) * "<row r=\"3\">$(bad_bool("B3"))<c r=\"C3\"><v>30</v></c></row>", String[])

    for (label, bytes) in ("outside the columns" => outside_column, "after the table" => after_table)
        t = XLSX.readtable(IOBuffer(bytes), 1, "B:C")
        @test t.column_labels == [:a, :b]
        @test t.data == Any[[2, 3], [20, 30]]
        XLSX._READTABLE_VALUE_ROWS[] = false
        try
            @test_throws XLSX.XLSXError XLSX.readtable(IOBuffer(bytes), 1, "B:C")
        finally
            XLSX._READTABLE_VALUE_ROWS[] = true
        end
    end
    # inside the read area: the same error on both paths
    e1 = try XLSX.readtable(IOBuffer(inside), 1, "B:C") catch e; e end
    XLSX._READTABLE_VALUE_ROWS[] = false
    e2 = try XLSX.readtable(IOBuffer(inside), 1, "B:C") catch e; e end
    XLSX._READTABLE_VALUE_ROWS[] = true
    @test e1 isa XLSX.XLSXError && e2 isa XLSX.XLSXError && e1.msg == e2.msg
end

# ── Filtered cell kernel vs the node decoder ─────────────────────────────────

# The cells, formulas and bounds of a sheet as read before `_read_cells` (each
# selected `<c>` decoded with `Cell(::LazyNode, …)` and then skipped), and whether
# its rows are in ascending order.
function _node_read_cells(ws, rows, cols)
    sel(s, x) = isnothing(s) || x in s
    wb = XLSX.get_workbook(ws)
    local_formulas = Dict{XLSX.SheetCellRef,XLSX.AbstractFormula}()
    cells = XLSX.Cell[]
    row_min = col_min = typemax(Int)
    row_max = col_max = 0
    in_rows = false
    ascending, last_row = true, 0
    c = XML.Cursor(XLSX._open_sheetdata(ws))
    while XML.next!(c) !== nothing
        d = XML.depth(c)
        d == 1 && continue
        if d == 2 && XML.nodetype(c) == XML.Element && XLSX.localname(c) == "row"
            row = parse(Int, XML.get(c, "r", nothing))
            ascending &= row > last_row; last_row = row
            row_min = min(row_min, row); row_max = max(row_max, row)
            in_rows = sel(rows, row)
        elseif d == 3 && XML.nodetype(c) == XML.Element && XLSX.localname(c) == "c"
            col = XLSX._ref_column_number(XML.get(c, "r", ""))
            col_min = min(col_min, col); col_max = max(col_max, col)
            if in_rows && sel(cols, col)
                push!(cells, XLSX.Cell(XML.LazyNode(c), ws, XLSX.get_sst_prefix(ws), local_formulas,
                                       XLSX.get_xlsxfile(ws).load_formulas))
            end
            XML.skip_element!(c)
        else
            XML.skip_element!(c)
        end
    end
    bounds = col_max > 0 ? XLSX.CellRange(XLSX.CellRef(row_min, col_min), XLSX.CellRef(row_max, col_max)) : nothing
    return cells, local_formulas, bounds, ascending
end

_formulas_of(ws) = filter(p -> first(p).sheet == ws.name, XLSX.get_workbook(ws).formulas)

# Cells the node decoder reads unusually: CDATA and repeated `<v>`, a second `<is>`,
# an `<f>` before `<is>`, comments and unknown children, empty and self-closing cells
# (also last in a row and last in the sheet), and a shared formula.
function _kernel_odd_cells_sheet(; pretty::Bool)
    rows = [
        "<row r=\"1\"><c r=\"A1\"><v><![CDATA[5]]></v></c><c r=\"B1\"><v>1</v><v>2</v></c>" *
            "<c r=\"C1\" t=\"inlineStr\"><is><t>a</t></is><is><t>b</t></is></c>" *
            "<c r=\"D1\" t=\"inlineStr\"><f>1</f><is><t>x</t></is></c>" *
            "<c r=\"E1\"><!-- note --><v>7</v><extLst><ext uri=\"u\"><y/></ext></extLst></c>" *
            "<c r=\"F1\" t=\"s\" s=\"1\" cm=\"1\"><v>0</v></c><c r=\"G1\" s=\"2\"/></row>",
        "<row r=\"2\"/>",
        "<row r=\"3\"><c r=\"B3\"><f t=\"shared\" ref=\"B3:B4\" si=\"0\">A1*2</f><v>10</v></c><c r=\"C3\"><v></v></c></row>",
        "<row r=\"4\"><c r=\"B4\"><f t=\"shared\" si=\"0\"/><v>4</v></c><c r=\"AB4\" t=\"str\"><f>\"q\"</f><v>q</v></c></row>",
        "<row r=\"6\"><c r=\"C6\"/></row>",
    ]
    xml = join(rows)
    # whitespace between rows, cells and a cell's children (never inside a value)
    pretty && (xml = replace(xml, (t => "\n  " * t for t in ("<row ", "</row>", "<c ", "<v>", "<f", "<is>", "<!--", "<extLst>"))...))
    return _diff_build_xlsx(xml, ["s0"])
end

# `(label, workbook bytes, sheet index)` of every sheet the kernel is checked on: the
# empty-row corpus, special and random sheets, the odd cells and the test/data worksheets.
function _kernel_sheets()
    sheets = Tuple{String,Vector{UInt8},Int}[]
    for len in 0:2, states in Iterators.product(ntuple(_ -> _DIFF_ROW_STATES, len)...)
        push!(sheets, ("corpus $(collect(states))", _diff_corpus_sheet(collect(states)), 1))
    end
    append!(sheets, [("special \"$k\"", v, 1) for (k, v) in _diff_special_sheets()])
    append!(sheets, [("random seed $s", _diff_random_sheet(Random.MersenneTwister(s)), 1) for s in 1:60])
    push!(sheets, ("odd cells", _kernel_odd_cells_sheet(pretty = false), 1))
    push!(sheets, ("odd cells, pretty-printed", _kernel_odd_cells_sheet(pretty = true), 1))
    for file in sort(readdir(data_directory))
        any(ext -> endswith(lowercase(file), ext), (".xlsx", ".xlsm", ".xltx", ".xltm")) || continue
        bytes = read(joinpath(data_directory, file))
        xf = try XLSX.readxlsx(IOBuffer(bytes)) catch; continue end
        wb = XLSX.get_workbook(xf)
        for (i, ws) in enumerate(wb.sheets)
            XLSX.is_chartsheet(wb, ws.name) || push!(sheets, ("$file sheet $i", bytes, i))
        end
    end
    return sheets
end

@testset "filtered cell kernel matches the node decoder" begin
    sheets = _kernel_sheets()
    selectors = (nothing, 1:0, 2:4, 3:3, 1:2:9, [1, 3, 7], [2, 28])
    mismatches = String[]
    n = 0
    for (label, bytes, sheet) in sheets, load_formulas in (true, false)
        openws() = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, false, false; load_formulas)[sheet]
        for rows in selectors, cols in selectors
            n += 1
            tag = "$label lf=$load_formulas rows=$rows cols=$cols"
            ws = openws()
            expected = _diff_run(() -> _node_read_cells(ws, rows, cols))
            ws = openws()
            result = _diff_run(() -> XLSX._read_cells(ws, rows, cols; track_bounds = true))
            if expected isa Exception || result isa Exception
                # an invalid cell in the selection: the same error
                (typeof(expected) == typeof(result) && sprint(showerror, expected) == sprint(showerror, result)) ||
                    push!(mismatches, "$tag: $expected vs $result")
                continue
            end
            want, want_formulas, want_bounds, ascending = expected
            got, bounds = result
            got == want || push!(mismatches, "$tag: cells differ")
            isequal(bounds, want_bounds) || push!(mismatches, "$tag: bounds $bounds vs $want_bounds")
            _formulas_of(ws) == want_formulas || push!(mismatches, "$tag: formulas differ")

            # rows in ascending order: the kernel may stop early without tracking bounds
            ascending || continue
            ws = openws()
            got, bounds = XLSX._read_cells(ws, rows, cols; track_bounds = false)
            got == want || push!(mismatches, "$tag, no bounds: cells differ")
            bounds === nothing || push!(mismatches, "$tag, no bounds: bounds $bounds")
        end
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test n > 10_000
end

@testset "filtered cell kernel: early stop and dimension scan" begin
    # Without tracking bounds, nothing after the last selected row is read: an invalid
    # cell there raises no error. With tracking, the pass reaches it.
    bad = "<row r=\"5\"><c r=\"B5\" t=\"b\"><v>2</v></c></row>"
    bytes = _diff_build_xlsx("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" * bad, String[])
    ws = XLSX.openxlsx(IOBuffer(bytes); enable_cache = false)[1]
    cells, bounds = XLSX._read_cells(ws, 1:1, nothing; track_bounds = false)
    @test length(cells) == 1 && XLSX.getdata(ws, only(cells)) == 1 && bounds === nothing
    cells, bounds = XLSX._read_cells(ws, 1:1, nothing; track_bounds = true)
    @test length(cells) == 1 && bounds == XLSX.CellRange("A1:B5")
    @test isempty(first(XLSX._read_cells(ws, Int[], nothing; track_bounds = false)))
    @test_throws XLSX.XLSXError XLSX._read_cells(ws, nothing, nothing; track_bounds = false)

    # The odd cells, compact and pretty-printed, decode to the same values.
    for pretty in (false, true)
        odd = XLSX.openxlsx(IOBuffer(_kernel_odd_cells_sheet(; pretty)); enable_cache = false)[1]
        cells, bounds = XLSX._read_cells(odd, nothing, nothing; track_bounds = true)
        @test bounds == XLSX.CellRange("A1:AB6")
        values = Dict(string(c.ref) => XLSX.getdata(odd, c) for c in cells)
        @test isequal(values, Dict("A1" => 5, "B1" => 2, "C1" => "a", "D1" => "x", "E1" => 7, "F1" => "s0",
                                   "G1" => missing, "B3" => 10, "C3" => missing, "B4" => 4, "AB4" => "q", "C6" => missing))
        @test [c.formula for c in cells] == [false, false, false, false, false, false, false, true, false, true, true, false]
        @test sort([string(k.cellref) for k in keys(_formulas_of(odd))]) == ["AB4", "B3", "B4"]
    end

    # `_scan_dimension` decodes nothing, so it reads past the invalid cell.
    @test XLSX._scan_dimension(ws) == XLSX.CellRange("A1:B5")
    @test XLSX._scan_dimension(XLSX.openxlsx(IOBuffer(_diff_build_xlsx("<row r=\"3\"/>", String[]));
                                             enable_cache = false)[1]) === nothing
end

# ── Uncached ranged reads vs the row decoders they replaced ──────────────────

# The uncached reads as they were before `_read_cells`: `getcell` matched the row with
# `match_rows`, `getdata` iterated rows until past the range, and non-contiguous
# selections read one cell at a time. `getcellrange` matched the range's rows with
# `match_rows`, which returned no rows after one absent from the file; the oracle
# iterates rows as `getcellrange` did with an empty cache.
function _old_getcell(ws, r, c)
    sheetrows = XLSX.match_rows(ws, [r])
    length(sheetrows) == 1 && return XLSX.getcell(sheetrows[1], c)
    return XLSX.EmptyCell(XLSX.CellRef(r, c))
end
function _old_getdata(ws, rng::XLSX.CellRange)::Array{Any,2}
    result = Array{Any,2}(undef, size(rng))
    fill!(result, missing)
    top, bottom = XLSX.row_number(rng.start), XLSX.row_number(rng.stop)
    for sheetrow in XLSX.eachrow(ws)
        if top <= sheetrow.row <= bottom
            for column in XLSX.column_number(rng.start):XLSX.column_number(rng.stop)
                cell = XLSX.getcell(sheetrow, column)
                if !isempty(cell)
                    (r, c) = XLSX.relative_cell_position(cell, rng)
                    result[r, c] = XLSX.getdata(ws, cell)
                end
            end
        end
        sheetrow.row > bottom && break
    end
    return result
end
function _old_getcellrange(ws, rng::XLSX.CellRange)::Array{XLSX.AbstractCell,2}
    result = Array{Any,2}(undef, size(rng))
    for ref in rng
        (r, c) = XLSX.relative_cell_position(ref, rng)
        result[r, c] = XLSX.EmptyCell(ref)
    end
    top, bottom = XLSX.row_number(rng.start), XLSX.row_number(rng.stop)
    for sheetrow in XLSX.eachrow(ws)
        if top <= sheetrow.row <= bottom
            for column in XLSX.column_number(rng.start):XLSX.column_number(rng.stop)
                cell = XLSX.getcell(sheetrow, column)
                (r, c) = XLSX.relative_cell_position(cell, rng)
                result[r, c] = cell
            end
        end
        sheetrow.row > bottom && break
    end
    return result
end
_old_getdata(ws, rows, cols) = [XLSX.getdata(ws, _old_getcell(ws, a, b)) for a in rows, b in cols]
_old_getcellrange(ws, rows, cols) = [_old_getcell(ws, a, b) for a in rows, b in cols]

# A cell's contents, independent of the order inline strings were added to the
# workbook's string table (which depends on how many cells a read decoded).
_cell_contents(ws, c::XLSX.Cell) = (c.ref, c.datatype, c.style, c.meta, c.formula, XLSX.getdata(ws, c))
_cell_contents(ws, x) = x
_contents(ws, x) = x isa AbstractArray ? map(c -> _cell_contents(ws, c), x) : _cell_contents(ws, x)

@testset "uncached ranged reads match the row decoders" begin
    ranges = XLSX.CellRange.(["A1:A1", "B2:D4", "A1:G9", "C3:AB6", "E7:H12", "B1:B30"])
    singles = XLSX.CellRef.(["A1", "B3", "C1", "AB4", "G7", "Z99"])
    selections = (([1, 3, 7], [2, 4]), (1:2:9, 2:3), (9:-2:1, [7, 2, 2]), (2, [1, 3]),
                  ([3, 1, 3], 2:4), (2:5, 1:3:7), (Int[], [1]), (1:1:4, 3))
    mismatches = String[]
    n = 0
    # Same result (contents and type), and formulas recorded for exactly the selected
    # cells; an error only where the old path raised one. (The old paths lost formulas:
    # the row iterator never merged those after its last 500-row batch.)
    # Like the other streaming reads, these assume ascending rows: a sheet with
    # out-of-order rows (never written by Excel) may read differently.
    function check(tag, openws, new, old, (rows, cols); malformed = false)
        n += 1
        ws_new, ws_old = openws(), openws()
        got, want = _diff_run(() -> new(ws_new)), _diff_run(() -> old(ws_old))
        if got isa Exception
            want isa Exception || push!(mismatches, "$tag: raised $got")
            return
        end
        want isa Exception && return  # the old path decoded an invalid cell outside the selection
        (malformed || isequal(_contents(ws_new, got), _contents(ws_old, want)) && typeof(got) == typeof(want)) ||
            push!(mismatches, "$tag: $(typeof(got)) $got vs $(typeof(want)) $want")
        _formulas_of(ws_new) == _node_read_cells(openws(), rows, cols)[2] ||
            push!(mismatches, "$tag: formulas differ")
    end
    span(rng) = (XLSX.row_number(rng.start):XLSX.row_number(rng.stop), XLSX.column_number(rng.start):XLSX.column_number(rng.stop))
    for (label, bytes, sheet) in _kernel_sheets(), load_formulas in (true, false)
        openws() = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, false, false; load_formulas)[sheet]
        tag = "$label lf=$load_formulas"
        malformed = label == "special \"out-of-order rows\""
        for rng in ranges
            check("$tag getdata $rng", openws, ws -> XLSX.getdata(ws, rng), ws -> _old_getdata(ws, rng), span(rng);
                  malformed)
            check("$tag getcellrange $rng", openws, ws -> XLSX.getcellrange(ws, rng), ws -> _old_getcellrange(ws, rng),
                  span(rng); malformed)
        end
        for ref in singles
            r, c = XLSX.row_number(ref), XLSX.column_number(ref)
            check("$tag getcell $ref", openws, ws -> XLSX.getcell(ws, ref), ws -> _old_getcell(ws, r, c), (r:r, c:c);
                  malformed)
        end
        for (rows, cols) in selections
            check("$tag getdata $rows×$cols", openws, ws -> XLSX.getdata(ws, rows, cols),
                  ws -> _old_getdata(ws, rows, cols), (rows, cols); malformed)
            check("$tag getcellrange $rows×$cols", openws, ws -> XLSX.getcellrange(ws, rows, cols),
                  ws -> _old_getcellrange(ws, rows, cols), (rows, cols); malformed)
        end
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test n > 10_000
end

@testset "uncached ranged reads: early stop and cached lookups" begin
    # Only the selected cells are decoded: an invalid cell outside them raises no error.
    bad = "<row r=\"5\"><c r=\"B5\" t=\"b\"><v>2</v></c></row>"
    rows = "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"C1\" t=\"b\"><v>2</v></c></row>" *
           "<row r=\"2\"><c r=\"A2\"><v>3</v></c></row>"
    ws = XLSX.openxlsx(IOBuffer(_diff_build_xlsx(rows * bad, String[])); enable_cache = false)[1]
    @test XLSX.getcell(ws, "A2").ref == XLSX.CellRef("A2") && XLSX.getdata(ws, XLSX.getcell(ws, "A2")) == 3
    @test XLSX.getcell(ws, "B2") == XLSX.EmptyCell(XLSX.CellRef("B2"))
    @test XLSX.getcell(ws, "A3") == XLSX.EmptyCell(XLSX.CellRef("A3"))
    @test XLSX.getcellrange(ws, "A1:A2") == [XLSX.getcell(ws, "A1"); XLSX.getcell(ws, "A2");;]
    @test isequal(XLSX.getdata(ws, [2, 1], 1:1), Any[3; 1;;])
    @test isequal(XLSX.getcellrange(ws, [2], [1, 2]), [XLSX.getcell(ws, "A2") XLSX.EmptyCell(XLSX.CellRef("B2"))])
    @test_throws XLSX.XLSXError XLSX.getcell(ws, "C1")
    @test_throws XLSX.XLSXError XLSX.getcell(ws, [1, 2], 1:3)
    # Rows after one absent from the file are read (`match_rows` lost them).
    gap = XLSX.openxlsx(joinpath(data_directory, "two_tables.xlsx"); enable_cache = false)[2]
    @test XLSX.getcellrange(gap, "A1:C3")[3, :] == [XLSX.getcell(gap, "A3"), XLSX.getcell(gap, "B3"), XLSX.getcell(gap, "C3")]
    @test all(c -> c isa XLSX.Cell, XLSX.getcellrange(gap, "A3:C3"))

    # The range read stops after its last row, so it leaves an unknown dimension unknown.
    @test isnothing(ws.dimension)
    @test isequal(XLSX.getdata(ws, "A1:A2"), Any[1; 3;;])
    @test isequal(XLSX.getdata(ws, "A2:B3"), Any[3 missing; missing missing])
    @test isnothing(ws.dimension)

    # With the cache enabled, non-contiguous selections are looked up cell by cell, as before.
    for file in ("general.xlsx", "customXml.xlsx")
        path = joinpath(data_directory, file)
        cached = XLSX.readxlsx(path)[1]
        uncached = XLSX.openxlsx(path; enable_cache = false)[1]
        rows, cols = [3, 1, 2], 1:2:5
        @test isequal(XLSX.getdata(cached, rows, cols), XLSX.getdata(uncached, rows, cols))
        @test XLSX.getcellrange(cached, rows, cols) == XLSX.getcellrange(uncached, rows, cols)
    end
end

# ── Single-cursor row stream vs the node walk it replaced ────────────────────

# The rows of a sheet as the stream iterator read them before (each `<row>` node's
# `r`/`ht` attributes and `Cell(::LazyNode, …)` per `<c>`, the last of a repeated column),
# and the number of shared-string cells.
function _old_stream_rows(ws)
    rows = Tuple{Int,Union{Nothing,Float64},Dict{Int,XLSX.Cell}}[]
    sst = 0
    for row in XML.eachchildnode(XLSX._open_sheetdata(ws))
        (XML.nodetype(row) == XML.Element && XLSX.localname(row) == "row") || continue
        r, ht = nothing, nothing
        for (k, v) in XML.eachattribute(row)
            k == "r" && (r = parse(Int, v))
            k == "ht" && (ht = parse(Float64, v))
        end
        isnothing(r) && throw(XLSX.XLSXError("Row without 'r' attribute in worksheet $(ws.name)."))
        cells = Dict{Int,XLSX.Cell}()
        for c in XML.eachchildnode(row)
            (XML.nodetype(c) == XML.Element && XLSX.localname(c) == "c") || continue
            cell = XLSX.Cell(c, ws, XLSX.get_sst_prefix(ws), Dict{XLSX.SheetCellRef,XLSX.AbstractFormula}(), false)
            sst += cell.datatype == XLSX.CT_STRING
            cells[XLSX.column_number(cell)] = cell
        end
        push!(rows, (r, ht, cells))
    end
    return rows, sst
end

# Rows by contents (see `_cell_contents`): a targeted read numbers inline strings differently.
_row_contents(ws, rows) = [(r, ht, Dict(k => _cell_contents(ws, v) for (k, v) in cells)) for (r, ht, cells) in rows]

# The rows the stream iterator yields, copied (it reuses one `Dict` per pass); stops
# after `limit` rows.
function _stream_rows(ws; limit = typemax(Int))
    rows = Tuple{Int,Union{Nothing,Float64},Dict{Int,XLSX.Cell}}[]
    for r in XLSX.eachrow(ws)
        push!(rows, (XLSX.row_number(r), r.ht, copy(r.rowcells)))
        length(rows) >= limit && break
    end
    return rows
end

@testset "row stream matches the node walk" begin
    mismatches = String[]
    n = 0
    for (label, bytes, sheet) in _kernel_sheets(), load_formulas in (true, false)
        openws() = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, false, false; load_formulas)[sheet]
        tag = "$label lf=$load_formulas"
        malformed = label in ("special \"out-of-order rows\"", "special \"duplicated row\"")
        n += 1
        oldws = openws()
        expected = _diff_run(() -> _old_stream_rows(oldws))
        ws = openws()
        sst0, dim0 = ws.sst_count, ws.dimension
        got = _diff_run(() -> _stream_rows(ws))
        if expected isa Exception || got isa Exception
            (typeof(expected) == typeof(got) && sprint(showerror, expected) == sprint(showerror, got)) ||
                push!(mismatches, "$tag: $expected vs $got")
            continue
        end
        want, sst = expected
        got == want || push!(mismatches, "$tag: rows differ")
        ws.sst_count == sst0 + sst || push!(mismatches, "$tag: sst_count $(ws.sst_count) vs $(sst0 + sst)")
        _, node_formulas, bounds, _ = _node_read_cells(openws(), nothing, nothing)
        isequal(ws.dimension, something(dim0, bounds, Some(nothing))) ||
            push!(mismatches, "$tag: dimension $(ws.dimension) vs $(something(dim0, bounds, Some(nothing)))")
        # every formula, which the old iterator lost after its last 500-row batch
        _formulas_of(ws) == node_formulas || push!(mismatches, "$tag: formulas differ")

        # a pass that stops early records the formulas of the rows it read
        if !malformed && length(want) > 1
            ws = openws()
            partial = _stream_rows(ws; limit = 2)
            partial == want[1:2] || push!(mismatches, "$tag, 2 rows: rows differ")
            _formulas_of(ws) == _node_read_cells(openws(), first.(partial), nothing)[2] ||
                push!(mismatches, "$tag, 2 rows: formulas differ")
        end

        # match_rows: the wanted rows present in the file, in ascending order
        malformed && continue
        present = first.(want)
        for wanted in ([1], [3, 1, 3], [2, 4, 5, 9], collect(1:12), [10_000], isempty(present) ? Int[] : [last(present)])
            n += 1
            ws = openws()
            matched = [(XLSX.row_number(r), r.ht, r.rowcells) for r in XLSX.match_rows(ws, wanted)]
            isequal(_row_contents(ws, matched), _row_contents(oldws, filter(r -> first(r) in wanted, want))) ||
                push!(mismatches, "$tag match_rows $wanted: rows differ")
            _formulas_of(ws) == _node_read_cells(openws(), first.(matched), nothing)[2] ||
                push!(mismatches, "$tag match_rows $wanted: formulas differ")
        end
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test n > 4_000
end

@testset "row stream: find_row and row heights" begin
    rows = "<row r=\"1\" ht=\"20.5\"><c r=\"A1\"><v>1</v></c></row><row r=\"3\" ht=\"x\"><c r=\"A3\"><v>3</v></c></row>" *
           "<row r=\"4\"><c r=\"A4\"><f>A1*4</f><v>4</v></c></row>"
    ws = XLSX.openxlsx(IOBuffer(_diff_build_xlsx(rows, String[])); enable_cache = false)[1]
    itr = XLSX.eachrow(ws)
    @test XLSX.find_row(itr, 1).ht == 20.5
    @test XLSX.find_row(itr, 4).ht === nothing && XLSX.getdata(ws, XLSX.getcell(XLSX.find_row(itr, 4), 1)) == 4
    @test_throws XLSX.XLSXError XLSX.find_row(itr, 2)
    @test_throws ArgumentError XLSX.find_row(itr, 3)                       # an invalid `ht` on a matched row
    @test XLSX.row_number.(XLSX.match_rows(ws, [1, 4])) == [1, 4]          # ... and not on a skipped one
    @test isempty(XLSX.match_rows(ws, Int[]))
    @test_throws ArgumentError collect(XLSX.eachrow(ws))
    @test haskey(XLSX.get_workbook(ws).formulas, XLSX.SheetCellRef(ws.name, XLSX.CellRef("A4")))
end

@testset "row stream: column window" begin
    mismatches = String[]
    n = 0
    for (label, bytes, sheet) in _kernel_sheets(), load_formulas in (true, false)
        openws() = XLSX.open_or_read_xlsx(IOBuffer(bytes), true, false, false; load_formulas)[sheet]
        full_ws = openws()
        full = _diff_run(() -> _stream_rows(full_ws))
        for cols in (1:1, 2:4, 3:3, 1:7, 28:28, 100:120)
            n += 1
            tag = "$label lf=$load_formulas cols=$cols"
            ws = openws()
            sst0 = ws.sst_count
            got = _diff_run(() -> begin
                rows = Tuple{Int,Union{Nothing,Float64},Dict{Int,XLSX.Cell}}[]
                for r in XLSX.SheetRowStreamIterator(ws, cols)
                    push!(rows, (XLSX.row_number(r), r.ht, copy(r.rowcells)))
                end
                rows
            end)
            if got isa Exception
                # an invalid cell inside the window raises as in the full pass
                full isa Exception || push!(mismatches, "$tag: raised $got")
                continue
            end
            full isa Exception && continue  # the full pass decoded an invalid cell outside the window
            want = [(r, ht, filter(p -> first(p) in cols, cells)) for (r, ht, cells) in full]
            isequal(_row_contents(ws, got), _row_contents(full_ws, want)) || push!(mismatches, "$tag: rows differ")
            ws.sst_count == sst0 + count(c -> c.datatype == XLSX.CT_STRING, (c for (_, _, cells) in want for c in values(cells))) ||
                push!(mismatches, "$tag: sst_count")
            isequal(ws.dimension, full_ws.dimension) || push!(mismatches, "$tag: dimension $(ws.dimension) vs $(full_ws.dimension)")
            _formulas_of(ws) == _node_read_cells(openws(), nothing, cols)[2] || push!(mismatches, "$tag: formulas differ")
        end
    end
    foreach(println, first(mismatches, 20))
    @test isempty(mismatches)
    @test n > 3_000

    # Uncached table reads decode only the table's columns: an invalid cell elsewhere
    # raises no error, one inside does.
    bad = "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>h</t></is></c></row>" *
          "<row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"C2\" t=\"b\"><v>2</v></c></row>" *
          "<row r=\"3\"><c r=\"A3\"><v>2</v></c></row>"
    open_bad() = XLSX.openxlsx(IOBuffer(_diff_build_xlsx(bad, String[])); enable_cache = false)[1]
    @test [r[:h] for r in XLSX.eachtablerow(open_bad(), "A:A")] == [1, 2]
    @test XLSX.gettable(open_bad(), "A"; first_row = 1).data == [[1, 2]]
    @test_throws XLSX.XLSXError XLSX.gettable(open_bad(), "A:C")

    # A `<c>` without `r` can't be placed in or out of the window.
    nor = XLSX.openxlsx(IOBuffer(_diff_build_xlsx("<row r=\"1\"><c><v>1</v></c></row>", String[])); enable_cache = false)[1]
    @test_throws XLSX.XLSXError collect(XLSX.SheetRowStreamIterator(nor, 2:3))
end
