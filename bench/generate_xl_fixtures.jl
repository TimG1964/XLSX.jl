# generate_xl_fixtures.jl
# Generates Excel-like .xlsm fixtures modelled on the workbook in issue #462, which
# can't be shared. Unlike generate_fixtures.jl (which writes through XLSX.jl and so
# produces pretty-printed XML), this writes the package parts directly, in the
# compact form Excel itself writes.
#
# Run once: julia --project=. generate_xl_fixtures.jl
#
# Inferred shape (see the #462 plan): a macro-enabled workbook with a small first
# sheet and a large second sheet; title rows 1–4, headers on row 5, the table in
# A:CF (84 columns) with 16 helper columns beyond it (CG:CV); `s` on every cell;
# numbers, shared-string categories, text, dates, booleans; 20% formula cells with
# cached values, half as shared formulas, some returning text and some #N/A;
# styled-but-empty cells throughout. In xl_medium and xl_issue462, rows empty within
# A:CF (helper cells only) and absent rows come only after the table: `readtable`
# stops at the first empty row by default, and the #462 call evidently read the whole
# sheet. xl_gaps also has gap rows inside the table (absent, helper-only and empty
# `<row/>`), for benchmarking reads with `stop_in_empty_row=false`. Every `<row>`
# carries `spans` and `x14ac:dyDescent`, as Excel writes them.
#
# The fixtures and their sizes are listed in XL_FIXTURES (xl_fixture_spec.jl).
#
# Every cell's content is a pure function of (row, column), so readers can check
# values without a stored copy: see `xl_expected` in xl_fixture_spec.jl.

using ZipArchives

include(joinpath(@__DIR__, "xl_fixture_spec.jl"))

const FIXTURES_DIR = get(ENV, "XL_FIXTURES_DIR", joinpath(@__DIR__, "fixtures"))

# Non-shared formulas as long lookups (true) or short arithmetic (false).
const LOOKUP_FORMULAS = Ref(true)
mkpath(FIXTURES_DIR)

xml_escape(s) = replace(s, "&" => "&amp;", "<" => "&lt;", ">" => "&gt;")

function write_cell(io, r, c, sst_index)
    ref = col_letter(c) * string(r)
    if styled_empty(r, c)
        print(io, "<c r=\"", ref, "\" s=\"6\"/>")
        return
    end
    k = KINDS[c]
    v = xl_expected(r, c)
    if k === :num
        print(io, "<c r=\"", ref, "\" s=\"3\"><v>", v, "</v></c>")
    elseif k === :int
        print(io, "<c r=\"", ref, "\" s=\"4\"><v>", v, "</v></c>")
    elseif k === :cat || k === :text
        print(io, "<c r=\"", ref, "\" s=\"5\" t=\"s\"><v>", sst_index(v), "</v></c>")
    elseif k === :date
        print(io, "<c r=\"", ref, "\" s=\"1\"><v>", v, "</v></c>")
    elseif k === :bool
        print(io, "<c r=\"", ref, "\" s=\"5\" t=\"b\"><v>", v ? 1 : 0, "</v></c>")
    elseif k === :fnum
        src = col_letter(c - 1) * string(r)
        if v isa String
            print(io, "<c r=\"", ref, "\" s=\"3\" t=\"e\"><f>NA()</f><v>#N/A</v></c>")
        else
            if LOOKUP_FORMULAS[]
                print(io, "<c r=\"", ref, "\" s=\"3\"><f>IFERROR(VLOOKUP(\$", col_letter(c - 2), r, ",Lookup!\$A\$2:\$H\$500,", 2 + c % 6, ",FALSE)*", src, "/20,0)</f><v>", v, "</v></c>")
            else
                print(io, "<c r=\"", ref, "\" s=\"3\"><f>ROUND(", src, "/20,2)</f><v>", v, "</v></c>")
            end
        end
    elseif k === :fshared
        si = c                                   # one shared group per column
        if r == FIRST_DATA
            print(io, "<c r=\"", ref, "\" s=\"3\"><f t=\"shared\" ref=\"", ref, ":", col_letter(c), LAST_ROW[], "\" si=\"", si,
                  "\">ROUND(", col_letter(c - 1), r, "/20,2)</f><v>", v, "</v></c>")
        else
            print(io, "<c r=\"", ref, "\" s=\"3\"><f t=\"shared\" si=\"", si, "\"/><v>", v, "</v></c>")
        end
    elseif k === :fstr
        print(io, "<c r=\"", ref, "\" s=\"5\" t=\"str\"><f>\"R\"&amp;ROW()</f><v>", v, "</v></c>")
    end
end

const SHEET_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0100-000000000000}">"""

function data_sheet_xml(nrows, sst_index)
    last = LAST_ROW[] = FIRST_DATA + nrows - 1
    io = IOBuffer()
    print(io, SHEET_HEAD, "<dimension ref=\"A1:", col_letter(NCOLS), last + TRAILER, "\"/>",
          "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"5\" topLeftCell=\"A6\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>",
          "<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>",
          "<cols><col min=\"1\" max=\"", NCOLS, "\" width=\"12.7109375\" customWidth=\"1\"/></cols><sheetData>")
    spans = "1:$NCOLS"
    # Title and notes rows 1–4 (row 2 absent, as an unused row would be).
    print(io, "<row r=\"1\" spans=\"", spans, "\" ht=\"21\" customHeight=\"1\" x14ac:dyDescent=\"0.35\"><c r=\"A1\" s=\"7\" t=\"s\"><v>", sst_index("Monthly operations extract"), "</v></c></row>")
    print(io, "<row r=\"3\" spans=\"", spans, "\" x14ac:dyDescent=\"0.25\"><c r=\"A3\" s=\"5\" t=\"s\"><v>", sst_index("Source: operations ledger"), "</v></c></row>")
    print(io, "<row r=\"4\" spans=\"", spans, "\" x14ac:dyDescent=\"0.25\"/>")
    print(io, "<row r=\"5\" spans=\"", spans, "\" x14ac:dyDescent=\"0.25\">")
    for c in 1:NCOLS
        label = c <= TABLE_COLS ? "$(KINDS[c])_$(c)" : "helper_$(c - TABLE_COLS)"
        print(io, "<c r=\"", col_letter(c), "5\" s=\"8\" t=\"s\"><v>", sst_index(label), "</v></c>")
    end
    print(io, "</row>")
    for r in FIRST_DATA:(last + TRAILER)
        cls = row_class(r)
        cls === :absent && continue
        if cls === :norow       # a formatted but empty row, as Excel writes one
            print(io, "<row r=\"", r, "\" spans=\"", spans, "\" x14ac:dyDescent=\"0.25\"/>")
            continue
        end
        print(io, "<row r=\"", r, "\" spans=\"", spans, "\" x14ac:dyDescent=\"0.25\">")
        for c in (cls === :helper_only ? (TABLE_COLS+1:NCOLS) : (1:NCOLS))
            write_cell(io, r, c, sst_index)
        end
        print(io, "</row>")
    end
    print(io, "</sheetData><mergeCells count=\"1\"><mergeCell ref=\"A1:F1\"/></mergeCells>",
          "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>")
    return take!(io)
end

function summary_sheet_xml(sst_index)
    io = IOBuffer()
    print(io, SHEET_HEAD, "<dimension ref=\"A1:B20\"/><sheetData>")
    for r in 1:20
        print(io, "<row r=\"", r, "\" spans=\"1:2\" x14ac:dyDescent=\"0.25\"><c r=\"A", r, "\" s=\"5\" t=\"s\"><v>",
              sst_index("Metric $r"), "</v></c><c r=\"B", r, "\" s=\"3\"><f>SUM(Data!A6:A100)</f><v>", r * 1.5, "</v></c></row>")
    end
    print(io, "</sheetData><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>")
    return take!(io)
end

const CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>"""

const ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""

const WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr codeName="ThisWorkbook" defaultThemeVersion="202300"/><bookViews><workbookView xWindow="-120" yWindow="-120" windowWidth="29040" windowHeight="15720" activeTab="1"/></bookViews><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="191029"/></workbook>"""

const WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"""

# cellXfs: 0 General, 1 date (14), 2 unused, 3 "0.00", 4 "0", 5 text cells, 6 fill only
# (styled-empty), 7 title, 8 header.
const STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><fonts count="3" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="16"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="0"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>"""

const THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2><a:accent1><a:srgbClr val="156082"/></a:accent1><a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3><a:accent4><a:srgbClr val="0F9ED5"/></a:accent4><a:accent5><a:srgbClr val="A02B93"/></a:accent5><a:accent6><a:srgbClr val="4EA72E"/></a:accent6><a:hlink><a:srgbClr val="467886"/></a:hlink><a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos Narrow"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"""

const CORE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>bench</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created></cp:coreProperties>"""

const APP = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>"""

function generate(label, nrows; gaps::Bool = false)
    path = joinpath(FIXTURES_DIR, "$(label).xlsm")
    isfile(path) && (println("Skipping $label (exists)"); return)
    IN_TABLE_GAPS[] = gaps
    println("Generating $label ($nrows data rows × $NCOLS cols$(gaps ? ", gaps in the table" : ""))…")

    strings = String[]
    index = Dict{String,Int}()
    sst_index(s) = get!(() -> (push!(strings, s); length(strings) - 1), index, s)

    summary = summary_sheet_xml(sst_index)
    data = data_sheet_xml(nrows, sst_index)

    sst = IOBuffer()
    print(sst, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"",
          length(strings), "\" uniqueCount=\"", length(strings), "\">")
    for s in strings
        print(sst, "<si><t>", xml_escape(s), "</t></si>")
    end
    print(sst, "</sst>")

    open(path, "w") do out
        ZipWriter(out) do w
            for (name, bytes) in (
                    "[Content_Types].xml" => codeunits(CONTENT_TYPES),
                    "_rels/.rels"         => codeunits(ROOT_RELS),
                    "docProps/core.xml"   => codeunits(CORE),
                    "docProps/app.xml"    => codeunits(APP),
                    "xl/workbook.xml"     => codeunits(WORKBOOK),
                    "xl/_rels/workbook.xml.rels" => codeunits(WORKBOOK_RELS),
                    "xl/styles.xml"       => codeunits(STYLES),
                    "xl/theme/theme1.xml" => codeunits(THEME),
                    "xl/sharedStrings.xml" => take!(sst),
                    "xl/worksheets/sheet1.xml" => summary,
                    "xl/worksheets/sheet2.xml" => data,
                    "xl/vbaProject.bin"   => UInt8[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, zeros(UInt8, 504)...])
                zip_newfile(w, name; compress = true)
                write(w, bytes)
                zip_commitfile(w)
            end
        end
    end
    println("  → written $path (sheet XML $(round(length(data) / 2^20, digits = 1)) MB, ",
            "file $(round(filesize(path) / 2^20, digits = 1)) MB, $(length(strings)) shared strings)")
end

if abspath(PROGRAM_FILE) == @__FILE__
    for (name, spec) in pairs(XL_FIXTURES)
        generate(String(name), spec.rows; spec.gaps)
    end
    println("Done.")
end
