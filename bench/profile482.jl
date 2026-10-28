# profile482.jl
# Profiles one benchmark of bench_worker.jl in the active environment, to locate where
# XML.jl `main` and v0.4.6 differ (#482). Run from bench/ in each environment, with the
# threads run_benchmarks.jl gives the worker:
#   julia --project=envs/v0_13_2        --threads=8 profile482.jl [fixture] [benchmark] [seconds]
#   julia --project=envs/v0_13_2_xmlmain --threads=8 profile482.jl [fixture] [benchmark] [seconds]
#
# fixture defaults to xl_issue462, benchmark to readtable (also open, writetable,
# open_readwrite), seconds to 10. Writes, to PROFILE_DIR (default results/profiles):
#   <tag>_flat.txt   frames by samples, with their callees, Julia frames only
#   <tag>_self.txt   frames by samples in the frame itself
#   <tag>_tree.txt   the call tree, to depth 40
#   <tag>.pb.gz      for `pprof`, if PProf loads (e.g. from the default environment)
# where <tag> is <env>_<fixture>_<benchmark>.
using XLSX, Profile, Printf, TOML

const FIXTURES_DIR = get(ENV, "XL_FIXTURES_DIR", joinpath(@__DIR__, "fixtures"))
const OUT_DIR      = get(ENV, "PROFILE_DIR", joinpath(@__DIR__, "results", "profiles"))

fixture   = get(ARGS, 1, "xl_issue462")
benchmark = get(ARGS, 2, "readtable")
seconds   = parse(Float64, get(ARGS, 3, "10"))

ext   = startswith(fixture, "xl_") ? ".xlsm" : ".xlsx"
path  = joinpath(FIXTURES_DIR, fixture * ext)
sheet = startswith(fixture, "xl_") ? "Data" : "Sheet1"
tmp   = tempname() * ".xlsx"

# As in bench_worker.jl
run_once = Dict(
    "open"           => () -> XLSX.openxlsx(_ -> nothing, path),
    "readtable"      => () -> XLSX.readtable(path, sheet),
    "writetable"     => () -> XLSX.writetable(tmp, XLSX.readtable(path, sheet); overwrite = true),
    "open_readwrite" => () -> begin
        tmp_rw = tempname() * ext
        cp(path, tmp_rw)
        XLSX.openxlsx(_ -> nothing, tmp_rw; mode = "rw")
        rm(tmp_rw; force = true)
    end,
)[benchmark]

env = basename(dirname(Base.active_project()))
tag = "$(env)_$(fixture)_$(benchmark)"
xml_rev = let m = joinpath(dirname(Base.active_project()), "Manifest.toml")
    x = only(TOML.parsefile(m)["deps"]["XML"])
    string(get(x, "repo-rev", "registry"), " ", first(get(x, "git-tree-sha1", "?"), 10))
end
header = "Julia $(VERSION) | XLSX $(pkgversion(XLSX)) | XML $(pkgversion(XLSX.XML)) $xml_rev | " *
         "threads $(Threads.nthreads()) | $fixture $benchmark"
println(header)

run_once(); run_once()                     # compile, and warm the file cache
times = Float64[]
for _ in 1:5
    GC.gc()
    push!(times, @elapsed run_once())
end
@printf("unprofiled: median %.1f ms of 5\n", 1e3 * sort(times)[3])

Profile.init(n = 10^8, delay = 0.001)
Profile.clear()
GC.gc()
nruns = 0
t0 = time()
@profile while time() - t0 < seconds
    run_once()
    global nruns += 1
end
println("profiled: $nruns runs in $(round(time() - t0; digits = 1)) s")

mkpath(OUT_DIR)
function save(suffix; kw...)
    file = joinpath(OUT_DIR, "$(tag)_$suffix.txt")
    open(file, "w") do io
        println(io, header, " | $nruns runs")
        Profile.print(IOContext(io, :displaysize => (1000, 400)); C = false, kw...)
    end
    println("wrote ", file)
end
save("flat"; format = :flat, sortedby = :count, mincount = 20)
save("self"; format = :flat, sortedby = :overhead, groupby = :none, mincount = 20)
save("tree"; format = :tree, maxdepth = 40, mincount = 50, noisefloor = 2)

try
    @eval using PProf
    file = joinpath(OUT_DIR, "$tag.pb.gz")
    Base.invokelatest(PProf.pprof; out = file, web = false)
    println("wrote ", file)
catch e
    println("no pprof file: ", sprint(showerror, e))
end
rm(tmp; force = true)
