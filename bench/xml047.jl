# xml047.jl
# The two reads #482 moves onto XML.jl's public API, timed in the active environment:
# `splitNode` on the xl_issue462 Data sheet, with and without a CR, and `sst_load!` on
# sst_unique, with LF and CR LF line ends. Median of five warmed `@timed` samples, or of
# XML047_SAMPLES.
#
# Run from bench/ in each environment to compare, on one machine in one session:
#   julia --project=envs/v0_13_2 xml047.jl
#   julia --project=envs/v0_13_2_xmlmain xml047.jl
#   julia --project=envs/dev_xmlmain xml047.jl
#
# Needs the fixtures from generate_xl_fixtures.jl, generate_fixtures.jl and
# generate_cr_fixtures.jl. Each line ends with a digest of the result, line ends
# normalised, which must match across environments.
using XLSX, Printf, TOML

const FIXTURES_DIR = get(ENV, "XL_FIXTURES_DIR", joinpath(@__DIR__, "fixtures"))
const NSAMPLES = parse(Int, get(ENV, "XML047_SAMPLES", "5"))  # odd, for the median

function samples(f)
    f()
    f()
    allocated = Int[]
    times = Float64[]
    for _ in 1:NSAMPLES
        GC.gc()
        result = @timed f()
        push!(allocated, result.bytes)
        push!(times, result.time)
    end
    return (seconds = sort(times)[cld(NSAMPLES, 2)], bytes = sort(allocated)[cld(NSAMPLES, 2)])
end

entry(file, name) = String(XLSX.ZipArchives.zip_readentry(
    XLSX.ZipArchives.ZipReader(read(joinpath(FIXTURES_DIR, file))), name))

digest(s::AbstractString) = string(hash(replace(s, "\r\n" => "\n")); base = 16)

function report(label, f, check)
    s = samples(f)
    @printf("%-28s %9.1f ms %10.1f MiB   %s\n", label, 1e3 * s.seconds, s.bytes / 2^20, digest(check(f())))
end

function sst_load(path)
    xf = XLSX.openxlsx(path; mode = "r", enable_cache = false)
    wb = XLSX.get_workbook(xf)
    return () -> begin
        sst = XLSX.get_sst(wb)
        sst.is_loaded = false
        XLSX.sst_load!(wb)
        sst
    end
end

xml = XLSX.XML
manifest = TOML.parsefile(joinpath(dirname(Base.active_project()), "Manifest.toml"))
xml_entry = only(manifest["deps"]["XML"])
println("Julia ", VERSION, " | XLSX ", pkgversion(XLSX), " | XML ", pkgversion(xml), " ",
        get(xml_entry, "repo-rev", "registry"), " ", first(get(xml_entry, "git-tree-sha1", "?"), 10),
        " | threads ", Threads.nthreads())

for (file, label) in (("xl_issue462.xlsm", "splitNode, no CR"), ("xl_issue462_cr.xlsm", "splitNode, one CR LF"))
    sheet = entry(file, "xl/worksheets/sheet2.xml")
    report(label, () -> first(XLSX.splitNode(sheet, "sheetData")), identity)
end
for (file, label) in (("sst_unique.xlsx", "LF"), ("sst_unique_crlf.xlsx", "CR LF"))
    path = joinpath(FIXTURES_DIR, file)
    report("sst_load!, $label", sst_load(path), sst -> join(sst.unformatted, '\0'))
end
