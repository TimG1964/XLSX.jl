# check_xl_fixtures.jl
# Reads each xl_* fixture as the benchmarks do and prints a digest of the result, so
# the same fixture can be compared across XLSX versions, and checks every value
# against the generator's spec (xl_fixture_spec.jl).
#
# xl_medium and xl_issue462 are read with the #462 call; xl_gaps with
# stop_in_empty_row=false, keep_empty_rows=true, so its gap rows come back as
# all-missing rows and table row i is always sheet row FIRST_DATA + i - 1.
#
# Usage: julia --project=envs/<ver> check_xl_fixtures.jl

using XLSX, Dates
include(joinpath(@__DIR__, "xl_fixture_spec.jl"))

# What XLSX.jl should return for a spec value.
function as_read(r, c)
    v = xl_expected(r, c)
    ismissing(v) && return missing
    KINDS[c] === :date && return Date(1899, 12, 30) + Day(v)
    return v
end

println("XLSX ", pkgversion(XLSX))
for (name, spec) in pairs(XL_FIXTURES)
    path = joinpath(@__DIR__, "fixtures", "$name.xlsm")
    isfile(path) || (println("  $name: missing"); continue)
    use_fixture!(name)
    kw = spec.gaps ? (; first_row = 5, stop_in_empty_row = false, keep_empty_rows = true) : (; first_row = 5)
    t = @elapsed (dt = XLSX.readtable(path, "Data", "A:CF"; kw...))
    data, labels = dt.data, dt.column_labels
    nrows = length(data[1])
    digest = hash((labels, map(eltype, data), [map(x -> ismissing(x) ? :missing : x, col) for col in data]))
    bad, errs, empty_rows = 0, 0, 0
    for i in 1:nrows
        r = FIRST_DATA + i - 1
        all(c -> ismissing(data[c][i]), 1:TABLE_COLS) && (empty_rows += 1)
        for c in 1:TABLE_COLS
            got, want = data[c][i], as_read(r, c)
            if want isa String && want == "#N/A"          # error cells: not compared
                errs += 1
                continue
            end
            ok = isequal(got, want) || (got isa AbstractFloat && want isa Real && isapprox(got, want))
            ok || (bad += 1; bad <= 5 && println("    mismatch row $r col $c: got $(repr(got)) want $(repr(want))"))
        end
    end
    # Rows the table should have: up to the last table row, or (reading past gaps)
    # up to the last row present in the sheet.
    want_rows = spec.gaps ? maximum(r for r in FIRST_DATA:(LAST_ROW[] + TRAILER) if row_class(r) !== :absent) - FIRST_DATA + 1 :
                            LAST_ROW[] - FIRST_DATA + 1
    nrows == want_rows || (bad += 1; println("    row count $nrows, expected $want_rows"))
    println("  $name: $(nrows) rows × $(length(labels)) cols, $empty_rows all-missing rows, digest $(string(digest, base = 16)), ",
            "value mismatches $bad, error cells $errs, eltypes $(sort(unique(string.(map(eltype, data)))))  [$(round(t, digits = 1)) s]")
end
