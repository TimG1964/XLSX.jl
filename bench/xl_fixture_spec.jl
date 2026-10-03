# xl_fixture_spec.jl
# The content of the xl_* fixtures as a pure function of (row, column), shared by
# generate_xl_fixtures.jl and check_xl_fixtures.jl. No package dependencies, so it
# loads in every benchmark environment.

const HEADER_ROW   = 5
const FIRST_DATA   = 6
const TABLE_COLS   = 84          # A:CF
const HELPER_COLS  = 16          # CG:CV
const NCOLS        = TABLE_COLS + HELPER_COLS
const LAST_ROW     = Ref(0)     # last table row, set per fixture
const TRAILER      = 30         # rows after the table: absent or helper-only
const IN_TABLE_GAPS = Ref(false) # gap rows inside the table too (xl_gaps)

# The fixtures: data rows, and whether the table itself has gap rows.
# `xl_gaps` is read with `stop_in_empty_row=false, keep_empty_rows=true`, so the gaps
# exercise the empty-row handling at scale; the others are read with the #462 call,
# which stops at the first empty row, so their gaps come only after the table.
const XL_FIXTURES = (
    xl_medium   = (rows = 3_000,  gaps = false),
    xl_issue462 = (rows = 30_000, gaps = false),
    xl_gaps     = (rows = 3_000,  gaps = true),
)

# Set LAST_ROW and IN_TABLE_GAPS for a fixture before generating or checking it.
function use_fixture!(name::Symbol)
    spec = XL_FIXTURES[name]
    LAST_ROW[] = FIRST_DATA + spec.rows - 1
    IN_TABLE_GAPS[] = spec.gaps
    return spec
end

col_letter(n) = n <= 26 ? string(Char('A' + n - 1)) : col_letter((n - 1) ÷ 26) * string(Char('A' + (n - 1) % 26))

# Deterministic per-cell hash in [0, 1).
h(r, c, salt = 0) = (hash((r, c, salt)) % 1_000_003) / 1_000_003

# Column kinds, fixed per column. Formula columns refer to the number column
# immediately to their left.
const KINDS = let k = Symbol[]
    pattern = (:num, :num, :cat, :num, :fnum, :date, :num, :text, :int, :fshared,
               :num, :cat, :bool, :fstr, :num, :date, :fshared, :num, :cat, :fnum)
    for c in 1:NCOLS
        push!(k, c > TABLE_COLS ? (isodd(c) ? :num : :cat) : pattern[mod1(c, length(pattern))])
    end
    k
end

const CATEGORIES = ["North", "South", "East", "West", "Central", "Retail", "Wholesale",
                    "Online", "Partner", "Direct", "Gold", "Silver", "Bronze", "Open",
                    "Closed", "Pending", "Approved", "Rejected", "Q1", "Q2", "Q3", "Q4"]

# Row classes. After the table, every third row is absent and the rest hold helper
# cells only (empty within A:CF). Inside the table, rows are all data unless
# IN_TABLE_GAPS is set: then 1% are absent, 1% helper-only and 0.5% an empty `<row/>`.
# The first data row is always data (it holds the shared formulas' master cells).
function row_class(r)
    r > LAST_ROW[] && return r % 3 == 0 ? :absent : :helper_only
    (IN_TABLE_GAPS[] && r > FIRST_DATA) || return :data
    x = h(r, 0, 1)
    return x < 0.01 ? :absent : x < 0.02 ? :helper_only : x < 0.025 ? :norow : :data
end
# 2% of data cells are styled-but-empty.
styled_empty(r, c) = h(r, c, 2) < 0.02

# Expected cell value as XLSX.jl reads it (for checking readers); `missing` for empty.
function xl_expected(r, c)
    (r < FIRST_DATA || row_class(r) in (:absent, :norow)) && return missing
    row_class(r) === :helper_only && c <= TABLE_COLS && return missing
    styled_empty(r, c) && return missing
    k = KINDS[c]
    k === :num     && return round(h(r, c) * 10_000, digits = 4)
    k === :int     && return Int(floor(h(r, c) * 1000))
    k === :cat     && return CATEGORIES[1 + Int(floor(h(r, c) * length(CATEGORIES)))]
    k === :text    && return "Item $(r)-$(c) $(Int(floor(h(r, c) * 1e6)))"
    k === :date    && return 40_000 + Int(floor(h(r, c) * 7000))       # serial; read as Date
    k === :bool    && return h(r, c) < 0.5
    k === :fnum    && return h(r, c, 3) < 0.01 ? "#N/A" : round(h(r, c) * 500, digits = 2)
    k === :fshared && return round(h(r, c) * 500, digits = 2)
    k === :fstr    && return "R$(r)"
    error("unknown kind $k")
end

