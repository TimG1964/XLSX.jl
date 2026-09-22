# chart_move_tools.jl — helpers for moving chart code into XLSX.Charts.
# Run from the package root, after the files are in src/charts/.
#
#   include("chart_move_tools.jl")
#   rename_refs()                # dry run: which files would change, and how
#   rename_refs(apply = true)    # rewrite `XLSX.name` → `XLSX.Charts.name`
#
#   using XLSX                   # once the module loads
#   check_imports()              # names Charts uses but can't see, or sees differently

const CHART_DIR = joinpath("src", "charts")
const EXCLUDE   = Set([:iserror, :geterror])   # core generics Charts only extends

# ── definitions (as in chart_inventory.jl, trimmed) ─────────────────────────────────────
function _signame(sig)
    while sig isa Expr && sig.head in (:where, :(::))
        sig = sig.args[1]
    end
    sig isa Expr && sig.head === :call && sig.args[1] isa Symbol ? sig.args[1] : nothing
end
function _typename(e)
    e isa Expr && e.head in (:(<:), :(::)) && (e = e.args[1])
    e isa Expr && e.head === :curly && (e = e.args[1])
    e isa Symbol ? e : nothing
end
function _defs!(out, ex)
    ex isa Expr || return out
    h = ex.head
    add(n) = n isa Symbol && push!(out, n)
    if h === :function && ex.args[1] isa Symbol
        add(ex.args[1]); return out
    elseif h === :function || h === :(=)
        n = _signame(ex.args[1])
        n === nothing || (add(n); return out)
        h === :function && return out
    elseif h === :struct
        add(_typename(ex.args[2])); return out
    elseif h === :abstract
        add(_typename(ex.args[1])); return out
    elseif h === :const && ex.args[1] isa Expr && ex.args[1].head === :(=)
        add(_typename(ex.args[1].args[1])); return out
    end
    foreach(a -> _defs!(out, a), ex.args)
    return out
end

jlfiles(dir) = [joinpath(d, f) for (d, _, fs) in walkdir(dir) for f in fs if endswith(f, ".jl")]
mdfiles(dir) = [joinpath(d, f) for (d, _, fs) in walkdir(dir) for f in fs if endswith(f, ".md")]
parsed(f)    = Meta.parseall(read(f, String); filename = f)

function chart_names()
    charts = Set{Symbol}()
    foreach(f -> _defs!(charts, parsed(f)), jlfiles(CHART_DIR))
    core = Set{Symbol}()
    for f in jlfiles("src")
        startswith(normpath(f), normpath(CHART_DIR)) || _defs!(core, parsed(f))
    end
    both = setdiff(intersect(charts, core), EXCLUDE)
    isempty(both) || @warn "Defined in both Charts and the core; not renamed" both
    return setdiff(charts, core, EXCLUDE)
end

# ── rename ───────────────────────────────────────────────────────────────────────────────
function rename_refs(; apply = false)
    names = sort!(string.(collect(chart_names())); by = length, rev = true)
    re = Regex("(?<![\\w.])XLSX\\.(" * join(names, "|") * ")(?![\\w!])")
    rewrite = vcat(jlfiles("test"), mdfiles(joinpath("docs", "src")), jlfiles(CHART_DIR))
    total = 0
    for f in rewrite
        txt = read(f, String)
        hits = unique(m.captures[1] for m in eachmatch(re, txt))
        isempty(hits) && continue
        n = count(re, txt)
        total += n
        println(rpad(f, 50), lpad(n, 5), "  ", join(first(hits, 6), ", "), length(hits) > 6 ? ", …" : "")
        apply && write(f, replace(txt, re => s"XLSX.Charts.\1"))
    end
    println(apply ? "Rewrote " : "Would rewrite ", total, " references.")
    # The core must not name Charts: report, never rewrite.
    for f in jlfiles("src")
        startswith(normpath(f), normpath(CHART_DIR)) && continue
        hits = unique(m.captures[1] for m in eachmatch(re, read(f, String)))
        isempty(hits) || @warn "Core file refers to chart names" f hits
    end
end

# ── import check ─────────────────────────────────────────────────────────────────────────
function _bare!(s, ex)
    if ex isa Symbol
        push!(s, ex)
    elseif ex isa Expr
        ex.head === :kw && length(ex.args) == 2 && return _bare!(s, ex.args[2])
        ex.head === :. && length(ex.args) == 2 && ex.args[2] isa QuoteNode &&
            return _bare!(s, ex.args[1])
        foreach(a -> _bare!(s, a), ex.args)
    end
    return s                                                  # QuoteNodes skipped
end

function check_imports()
    isdefined(Main, :XLSX) || error("run `using XLSX` first")
    X = Main.XLSX
    isdefined(X, :Charts) || error("XLSX.Charts is not defined yet")
    C = X.Charts
    own = chart_names()
    unseen   = Dict{Symbol,Set{String}}()
    shadowed = Dict{Symbol,Set{String}}()
    for f in jlfiles(CHART_DIR), s in _bare!(Set{Symbol}(), parsed(f))
        isdefined(X, s) || continue
        if !isdefined(C, s)
            push!(get!(Set{String}, unseen, s), basename(f))
        elseif getglobal(C, s) !== getglobal(X, s) && !(s in own)
            push!(get!(Set{String}, shadowed, s), basename(f))   # e.g. Base.f in Charts, XLSX.f in core
        end
    end
    println("Defined in XLSX but not visible in Charts (import, or a local of the same name):")
    for (s, fs) in sort!(collect(unseen); by = first)
        println("  ", rpad(s, 40), join(sort!(collect(fs)), ", "))
    end
    println("\nVisible in Charts but a different object from XLSX's:")
    for (s, fs) in sort!(collect(shadowed); by = first)
        println("  ", rpad(s, 40), join(sort!(collect(fs)), ", "))
    end
end
