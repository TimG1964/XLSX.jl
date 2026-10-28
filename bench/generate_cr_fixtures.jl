# generate_cr_fixtures.jl
# Copies of existing fixtures whose worksheet and shared-string XML has CR line ends,
# for #482: XML.jl copies a String that holds a CR to normalise its line ends, unless
# it is read through `XML.sourceview`.
#
#   xl_issue462_cr   xl_issue462 with one CR LF, after each part's XML declaration
#   large_crlf       large with every LF a CR LF
#   sst_unique_crlf  sst_unique with every LF a CR LF
#
# Run after generate_fixtures.jl and generate_xl_fixtures.jl:
#   julia --project=. generate_cr_fixtures.jl

using ZipArchives

const FIXTURES_DIR = get(ENV, "XL_FIXTURES_DIR", joinpath(@__DIR__, "fixtures"))

is_rewritten(name) = occursin(r"^xl/worksheets/[^/]+\.xml$", name) || name == "xl/sharedStrings.xml"

function rewrite(src_name, dst_name, f)
    src = joinpath(FIXTURES_DIR, src_name)
    dst = joinpath(FIXTURES_DIR, dst_name)
    isfile(dst) && (println("Skipping $dst_name (exists)"); return)
    isfile(src) || (println("Skipping $dst_name: no $src_name, run the other generators first"); return)
    println("Generating $dst_name from $src_name…")
    r = ZipReader(read(src))
    ncr = 0
    open(dst, "w") do out
        ZipWriter(out) do w
            for name in zip_names(r)
                bytes = zip_readentry(r, name)
                if is_rewritten(name)
                    bytes = Vector{UInt8}(f(String(bytes)))
                    ncr += count(==(0x0d), bytes)
                end
                zip_newfile(w, name; compress = true)
                write(w, bytes)
                zip_commitfile(w)
            end
        end
    end
    println("  → written $dst ($ncr CRs, file $(round(filesize(dst) / 2^20, digits = 1)) MB)")
end

rewrite("xl_issue462.xlsm", "xl_issue462_cr.xlsm", s -> replace(s, "?>\n" => "?>\r\n"; count = 1))
rewrite("large.xlsx",       "large_crlf.xlsx",      s -> replace(s, "\n" => "\r\n"))
rewrite("sst_unique.xlsx",  "sst_unique_crlf.xlsx", s -> replace(s, "\n" => "\r\n"))
