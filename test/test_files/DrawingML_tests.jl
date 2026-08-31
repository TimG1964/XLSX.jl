@testset "DrawingML colours" begin

    # A workbook is needed only for schemeClr resolution; any file with a
    # standard theme will do.
    f = XLSX.opentemplate(joinpath(data_directory, "chart_basic.xlsx"))
    wb = XLSX.get_workbook(f)

    # Parse a colour from an XML fragment, as it appears inside a fill.
    function color_of(xml::String)
        doc = parse(xml, XML.Node)
        XLSX.parse_drawing_color(wb, XLSX.xml_root_element(doc))
    end

    @testset "theme colour map" begin
        m = XLSX.get_theme_color_map(wb)
        for name in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                     "accent4", "accent5", "accent6", "hlink", "folHlink")
            @test haskey(m, name)
            @test occursin(r"^[0-9A-Fa-f]{6}$", m[name])
        end

        # DrawingML aliases, not index-swapped like the spreadsheet ordering.
        @test m["tx1"] == m["dk1"]
        @test m["bg1"] == m["lt1"]
        @test m["tx2"] == m["dk2"]
        @test m["bg2"] == m["lt2"]

        # Cached: the same dict object comes back.
        @test XLSX.get_theme_color_map(wb) === m
    end

    @testset "srgbClr" begin
        c = color_of("""<a:srgbClr xmlns:a="$(XLSX.NS_A)" val="FF0000"/>""")
        @test c.kind === :srgb
        @test c.val == "FF0000"
        @test c.rgb == "FF0000"
        @test c.alpha == 1.0
        @test isempty(c.transforms)
    end

    @testset "schemeClr resolves through the theme" begin
        c = color_of("""<a:schemeClr xmlns:a="$(XLSX.NS_A)" val="accent1"/>""")
        @test c.kind === :scheme
        @test c.val == "accent1"                       # the reference is kept as authored
        @test c.rgb == uppercase(XLSX.get_theme_color_map(wb)["accent1"])
        @test isempty(c.transforms)
    end

    @testset "lumMod and lumOff apply in document order" begin
        # "Accent 1, Lighter 40%" as Excel writes it.
        c = color_of("""
            <a:schemeClr xmlns:a="$(XLSX.NS_A)" val="accent1">
                <a:lumMod val="60000"/><a:lumOff val="40000"/>
            </a:schemeClr>""")
        @test c.transforms == [:lumMod => 60000, :lumOff => 40000]
        @test c.rgb != uppercase(XLSX.get_theme_color_map(wb)["accent1"])  # it moved
        @test occursin(r"^[0-9A-F]{6}$", c.rgb)

        # Lightening raises luminance; darkening lowers it.
        base  = color_of("""<a:schemeClr xmlns:a="$(XLSX.NS_A)" val="accent1"/>""")
        dark  = color_of("""
            <a:schemeClr xmlns:a="$(XLSX.NS_A)" val="accent1">
                <a:lumMod val="75000"/>
            </a:schemeClr>""")
        lum(h) = convert(Colors.HSL, parse(Colors.Colorant, "#" * h)).l
        @test lum(c.rgb) > lum(base.rgb)
        @test lum(dark.rgb) < lum(base.rgb)
    end

    @testset "alpha sets alpha, not rgb" begin
        c = color_of("""
            <a:srgbClr xmlns:a="$(XLSX.NS_A)" val="00FF00">
                <a:alpha val="50000"/>
            </a:srgbClr>""")
        @test c.rgb == "00FF00"
        @test c.alpha ≈ 0.5
    end

    @testset "shade and tint go through the gamma transfer" begin
        # A 50% shade of mid grey is not 50% of its sRGB value: applying the
        # transform in gamma-encoded space would give 3F3F3F.
        c = color_of("""
            <a:srgbClr xmlns:a="$(XLSX.NS_A)" val="7F7F7F">
                <a:shade val="50000"/>
            </a:srgbClr>""")
        @test c.rgb != "3F3F3F"
        @test parse(Int, c.rgb[1:2]; base=16) > 0x3F      # lighter than the naive result

        t = color_of("""
            <a:srgbClr xmlns:a="$(XLSX.NS_A)" val="000000">
                <a:tint val="50000"/>
            </a:srgbClr>""")
        @test t.rgb != "000000"                            # tint moves black toward white
    end

    @testset "sysClr prefers lastClr" begin
        c = color_of("""<a:sysClr xmlns:a="$(XLSX.NS_A)" val="windowText" lastClr="112233"/>""")
        @test c.kind === :sys
        @test c.val == "windowText"
        @test c.rgb == "112233"
    end

    @testset "unknown transforms are no-ops" begin
        plain = color_of("""<a:srgbClr xmlns:a="$(XLSX.NS_A)" val="336699"/>""")
        c = color_of("""
            <a:srgbClr xmlns:a="$(XLSX.NS_A)" val="336699">
                <a:hueMod val="50000"/>
            </a:srgbClr>""")
        @test c.rgb == plain.rgb                # not applied...
        @test c.transforms == [:hueMod => 50000]  # ...but recorded
    end

    @testset "no colour child" begin
        doc = parse("""<a:noFill xmlns:a="$(XLSX.NS_A)"/>""", XML.Node)
        @test isnothing(XLSX.parse_drawing_color(wb, XLSX.xml_root_element(doc)))
        @test isnothing(XLSX.parse_drawing_color(wb, nothing))
    end

    @testset "parent element is searched for its colour child" begin
        doc = parse("""
            <a:solidFill xmlns:a="$(XLSX.NS_A)"><a:srgbClr val="ABCDEF"/></a:solidFill>""", XML.Node)
        c = XLSX.parse_drawing_color(wb, XLSX.xml_root_element(doc))
        @test c.rgb == "ABCDEF"
    end

    fill_of(xml) = XLSX.parse_drawing_fill(wb, XLSX.xml_root_element(XML.parse(xml, XML.Node)))

    @testset "solid fill" begin
        fl = fill_of("""
            <a:solidFill xmlns:a="$(XLSX.NS_A)"><a:srgbClr val="FF0000"/></a:solidFill>""")
        @test fl.kind === :solid
        @test fl.fgcolor.rgb == "FF0000"
        @test isnothing(fl.bgcolor)
        @test isnothing(fl.preset)
    end

    @testset "no fill" begin
        fl = fill_of("""<a:noFill xmlns:a="$(XLSX.NS_A)"/>""")
        @test fl.kind === :none
        @test isnothing(fl.fgcolor)
    end

    @testset "pattern fill exposes both colours" begin
        fl = fill_of("""
            <a:pattFill xmlns:a="$(XLSX.NS_A)" prst="pct25">
                <a:fgClr><a:srgbClr val="112233"/></a:fgClr>
                <a:bgClr><a:schemeClr val="bg1"/></a:bgClr>
            </a:pattFill>""")
        @test fl.kind === :pattern
        @test fl.preset == "pct25"
        @test fl.fgcolor.rgb == "112233"
        @test fl.bgcolor.val == "bg1"
        @test fl.bgcolor.rgb == uppercase(XLSX.get_theme_color_map(wb)["bg1"])
    end

    @testset "gradient is identified but not modelled" begin
        fl = fill_of("""
            <a:gradFill xmlns:a="$(XLSX.NS_A)">
                <a:gsLst>
                    <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
                    <a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>
                </a:gsLst>
            </a:gradFill>""")
        @test fl.kind === :gradient
        @test isnothing(fl.fgcolor)          # not the first stop, which would mislead
        @test XLSX.localname(fl.raw) == "gradFill"
    end

    @testset "parent element is searched for its fill child" begin
        fl = fill_of("""
            <c:spPr xmlns:c="$(XLSX.NS_C)" xmlns:a="$(XLSX.NS_A)">
                <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
            </c:spPr>""")
        @test fl.fgcolor.rgb == "ABCDEF"
    end

    @testset "no fill child" begin
        doc = XML.parse("""<c:spPr xmlns:c="$(XLSX.NS_C)"/>""", XML.Node)
        @test isnothing(XLSX.parse_drawing_fill(wb, XLSX.xml_root_element(doc)))
        @test isnothing(XLSX.parse_drawing_fill(wb, nothing))
    end

    @testset "parent element is searched for its fill child" begin
        fl = fill_of("""
            <a:spPr xmlns:a="$(XLSX.NS_A)">
                <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
            </a:spPr>""")
        @test fl.fgcolor.rgb == "ABCDEF"
    end

    @testset "no fill child" begin
        doc = XML.parse("""<a:spPr xmlns:a="$(XLSX.NS_A)"/>""", XML.Node)
        @test isnothing(XLSX.parse_drawing_fill(wb, XLSX.xml_root_element(doc)))
        @test isnothing(XLSX.parse_drawing_fill(wb, nothing))
    end

    line_of(xml) = XLSX.parse_drawing_line(wb, XLSX.xml_root_element(XML.parse(xml, XML.Node)))

    @testset "solid line with width and dash" begin
        ln = line_of("""
            <a:ln xmlns:a="$(XLSX.NS_A)" w="19050" cap="rnd" cmpd="sng">
                <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                <a:prstDash val="dash"/>
            </a:ln>""")
        @test ln.fill.kind === :solid
        @test ln.fill.fgcolor.rgb == "FF0000"
        @test ln.width == 1.5
        @test ln.dash == "dash"          # an element, not an attribute
        @test ln.cap == "rnd"
        @test ln.compound == "sng"
    end

    @testset "line with no outline" begin
        ln = line_of("""
            <a:ln xmlns:a="$(XLSX.NS_A)"><a:noFill/></a:ln>""")
        @test ln.fill.kind === :none
        @test isnothing(ln.fill.fgcolor)
        @test isnothing(ln.width)
        @test isnothing(ln.dash)
    end

    @testset "line that sets only a width" begin
        ln = line_of("""<a:ln xmlns:a="$(XLSX.NS_A)" w="9525"/>""")
        @test isnothing(ln.fill)         # says nothing about the stroke...
        @test ln.width == 0.75
        @test isnothing(ln.cap)
        @test isnothing(ln.compound)
    end

    @testset "line colour goes through the theme" begin
        ln = line_of("""
            <a:ln xmlns:a="$(XLSX.NS_A)">
                <a:solidFill>
                    <a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr>
                </a:solidFill>
            </a:ln>""")
        @test ln.fill.fgcolor.val == "tx1"
        @test ln.fill.fgcolor.rgb == "D9D9D9"    # the standard Office gridline grey
    end

    @testset "parent element is searched for its ln child" begin
        ln = line_of("""
            <c:spPr xmlns:c="$(XLSX.NS_C)" xmlns:a="$(XLSX.NS_A)">
                <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
                <a:ln w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln>
            </c:spPr>""")
        # The line's own fill, not the shape's.
        @test ln.fill.fgcolor.rgb == "123456"
        @test ln.width == 1.0
    end

    @testset "no ln child" begin
        doc = XML.parse("""
            <c:spPr xmlns:c="$(XLSX.NS_C)" xmlns:a="$(XLSX.NS_A)">
                <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
            </c:spPr>""", XML.Node)
        @test isnothing(XLSX.parse_drawing_line(wb, XLSX.xml_root_element(doc)))
        @test isnothing(XLSX.parse_drawing_line(wb, nothing))
        ln_bare = line_of("""<a:ln xmlns:a="$(XLSX.NS_A)"/>""")
        @test ln_bare !== nothing
        @test isnothing(ln_bare.width)
    end
end

# =============================================================================
# DrawingML Text
#
# Three groups:
#   1. attribute readers — units and the absent/explicit distinction, no file
#   2. hand-built XML — shapes the fixtures may not contain (multi-run, mixed
#      formatting, breaks, indentation)
#   3. fixtures — the real thing, against chart_basic.xlsx and
#      chart_theme_colors.xlsx
#
# Group 3 has @test_broken-style placeholders marked TODO where I don't know
# what your fixtures contain. Run them, read the failures, pin the values.
# =============================================================================

const localname = XLSX.localname
const _attr           = XLSX._attr
const _attr_int       = XLSX._attr_int
const _attr_bool      = XLSX._attr_bool
const _attr_pt        = XLSX._attr_pt
const _attr_emu       = XLSX._attr_emu
const _attr_deg       = XLSX._attr_deg
const _attr_pct_opt   = XLSX._attr_pct_opt
const _attr_spacing   = XLSX._attr_spacing
const first_element_with_tag = XLSX.first_element_with_tag
const NS_C = XLSX.NS_C
const NS_A = XLSX.NS_A

const parse_drawing_text  = XLSX.parse_drawing_text
const default_run_props   = XLSX.default_run_props
const first_run_props     = XLSX.first_run_props
const is_uniform          = XLSX.is_uniform
const text_content        = XLSX.text_content
const text_runs           = XLSX.text_runs
const DrawingText         = XLSX.DrawingText
const get_theme_fonts     = XLSX.get_theme_fonts
const resolve_theme_font  = XLSX.resolve_theme_font
const xml_root_element    = XLSX.xml_root_element
const get_xml_data        = XLSX.get_xml_data
const getCharts           = XLSX.getCharts

const parse_drawing_shape_props = XLSX.parse_drawing_shape_props
const DrawingShapeProps         = XLSX.DrawingShapeProps
const has_fill                  = XLSX.has_fill
const has_line                  = XLSX.has_line

# Parse a fragment into its root element. XML.parse returns a document node;
# the element is its last child (a declaration may precede it).
function _frag(s::String)
    doc = XML.parse(XML.Node, s)
    els = filter(c -> XML.nodetype(c) === XML.Element, XML.children(doc))
    return els[end]
end

const _NSDECL = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""

@testset "DrawingML text" begin

    # -----------------------------------------------------------------------
    @testset "attribute readers" begin
        el = _frag("""<a:rPr $_NSDECL sz="1197" b="0" i="1" baseline="30000"
                      kern="1200" spc="-50" lvl="2" rot="-2700000" marL="228600"/>""")

        # units
        @test _attr_pt(el, "sz") == 11.97           # 1/100 pt
        @test _attr_pt(el, "kern") == 12.0
        @test _attr_pt(el, "spc") == -0.5
        @test _attr_pct_opt(el, "baseline") == 0.3  # thousandths of a percent -> fraction
        @test _attr_deg(el, "rot") == -45.0         # 1/60000 deg
        @test _attr_emu(el, "marL") == 18.0         # EMU -> pt
        @test _attr_int(el, "lvl") == 2

        # absent stays distinct from explicit
        @test _attr_bool(el, "b") === false         # written as 0
        @test _attr_bool(el, "i") === true
        @test _attr_bool(el, "u") === nothing       # not written at all
        @test _attr(el, "missing") === nothing
        @test _attr_pt(el, "missing") === nothing
        @test _attr_pct_opt(el, "missing") === nothing

        # tolerates a missing child element
        @test _attr(nothing, "typeface") === nothing
        @test first_element_with_tag(nothing, "latin") === nothing

        # bool spellings
        b2 = _frag("""<a:rPr $_NSDECL b="true" i="on" u="none"/>""")
        @test _attr_bool(b2, "b") === true
        @test _attr_bool(b2, "i") === true
        @test _attr(b2, "u") == "none"              # enum stays a string

        # unparseable reads as absent, not as zero
        bad = _frag("""<a:rPr $_NSDECL sz="large"/>""")
        @test _attr_pt(bad, "sz") === nothing
    end

    # -----------------------------------------------------------------------
    @testset "spacing" begin
        pPr = _frag("""<a:pPr $_NSDECL>
                         <a:lnSpc><a:spcPct val="150000"/></a:lnSpc>
                         <a:spcBef><a:spcPts val="1200"/></a:spcBef>
                       </a:pPr>""")
        @test _attr_spacing(pPr, "lnSpc") == (:frac, 1.5)
        @test _attr_spacing(pPr, "spcBef") == (:pts, 12.0)
        @test _attr_spacing(pPr, "spcAft") === nothing
    end

    # -----------------------------------------------------------------------
    @testset "hand-built text bodies" begin
        # A workbook is needed only for theme colour resolution; any will do.
        XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx")) do xf
            wb = XLSX.get_workbook(xf)

            @testset "formatting-only txPr (the common shape)" begin
                node = _frag("""<c:txPr xmlns:c="$NS_C" $_NSDECL>
                      <a:bodyPr rot="-60000000" vert="horz" wrap="square" anchor="ctr"/>
                      <a:lstStyle/>
                      <a:p>
                        <a:pPr algn="ctr">
                          <a:defRPr sz="900" b="1" i="0" u="none" strike="noStrike"/>
                        </a:pPr>
                        <a:endParaRPr lang="en-GB"/>
                      </a:p>
                    </c:txPr>""")

                t = parse_drawing_text(wb, node)
                @test t isa DrawingText
                @test length(t.paragraphs) == 1
                @test isempty(t.paragraphs[1].runs)
                @test text_content(t) == ""
                @test t.liststyle !== nothing              # preserved, even empty

                @test t.body.rotation == -1000.0
                @test t.body.anchor == "ctr"
                @test t.body.autofit === nothing           # absent, not :none

                # the font lives in defRPr, and default_run_props finds it
                p = default_run_props(t)
                @test p !== nothing
                @test p.size == 9.0
                @test p.bold === true
                @test p.italic === false
                @test p.under == "none"
                @test t.paragraphs[1].props.align == "ctr"

                @test is_uniform(t)                        # no runs: uniform
            end

            @testset "uniform multi-run rich text" begin
                node = _frag("""<c:rich xmlns:c="$NS_C" $_NSDECL>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr lang="en-GB" sz="1400" b="1"/><a:t>Quarterly </a:t></a:r>
                        <a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Revenue</a:t></a:r>
                      </a:p>
                    </c:rich>""")

                t = parse_drawing_text(wb, node; tag="rich")
                @test length(text_runs(t)) == 2
                @test text_content(t) == "Quarterly Revenue"

                # differing only by lang — Excel splits runs on spellcheck
                # boundaries, and that must not read as mixed formatting
                @test is_uniform(t)
                @test default_run_props(t) !== nothing
                @test default_run_props(t).size == 14.0
            end

            @testset "mixed formatting" begin
                node = _frag("""<c:rich xmlns:c="$NS_C" $_NSDECL>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="1400" b="0"/><a:t>Revenue </a:t></a:r>
                        <a:r><a:rPr sz="1400" b="1"/><a:t>2024</a:t></a:r>
                      </a:p>
                    </c:rich>""")

                t = parse_drawing_text(wb, node; tag="rich")
                @test text_content(t) == "Revenue 2024"
                @test !is_uniform(t)

                # no single answer, so no answer
                @test default_run_props(t) === nothing
                # ...but the first fragment is still reachable
                @test first_run_props(t).bold === false
                @test text_runs(t)[2].props.bold === true
            end

            @testset "breaks and paragraphs" begin
                node = _frag("""<c:rich xmlns:c="$NS_C" $_NSDECL>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:t>One</a:t></a:r><a:br/><a:r><a:t>Two</a:t></a:r></a:p>
                      <a:p><a:r><a:t>Three</a:t></a:r></a:p>
                    </c:rich>""")

                t = parse_drawing_text(wb, node; tag="rich")
                @test length(t.paragraphs) == 2
                @test text_content(t) == "One\nTwo\nThree"
                @test [r.kind for r in t.paragraphs[1].runs] == [:run, :br, :run]
            end

            @testset "indented XML (nodetype guard)" begin
                # Excel writes chart parts unindented; a formatted or
                # hand-edited part has whitespace text nodes between runs.
                node = _frag("""<c:rich xmlns:c="$NS_C" $_NSDECL>
                      <a:bodyPr/>
                      <a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="1000"/><a:t>Spaced</a:t></a:r>
                        <a:r><a:rPr sz="1000"/><a:t> out</a:t></a:r>
                      </a:p>
                    </c:rich>""")

                t = parse_drawing_text(wb, node; tag="rich")
                @test length(text_runs(t)) == 2          # not 2 + whitespace nodes
                @test text_content(t) == "Spaced out"
            end

            @testset "parse from parent, and absent txPr" begin
                parent = _frag("""<c:valAx xmlns:c="$NS_C" $_NSDECL>
                      <c:delete val="0"/>
                      <c:txPr><a:bodyPr/><a:lstStyle/>
                        <a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p>
                      </c:txPr>
                    </c:valAx>""")
                t = parse_drawing_text(wb, parent)          # searches for txPr
                @test t !== nothing
                @test default_run_props(t).size == 10.0

                bare = _frag("""<c:valAx xmlns:c="$NS_C"><c:delete val="0"/></c:valAx>""")
                @test parse_drawing_text(wb, bare) === nothing
            end

            @testset "solidFill inside defRPr" begin
                node = _frag("""<c:txPr xmlns:c="$NS_C" $_NSDECL>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:defRPr sz="900">
                        <a:solidFill><a:schemeClr val="tx1">
                          <a:lumMod val="65000"/><a:lumOff val="35000"/>
                        </a:schemeClr></a:solidFill>
                        <a:latin typeface="+mn-lt"/>
                      </a:defRPr></a:pPr></a:p>
                    </c:txPr>""")

                t = parse_drawing_text(wb, node)
                p = default_run_props(t)
                @test p.fill !== nothing
                @test p.fill.kind == :solid
                # same transform the colour tests already pin
                @test p.fill.fgcolor.rgb == "595959"
                @test p.latin == "+mn-lt"           # kept as written
                @test p.ea === nothing
                @test p.line === nothing
            end
        end
    end

    # -----------------------------------------------------------------------
        # -----------------------------------------------------------------------
    @testset "fixtures" begin

        # NOTE: `Chart` is metadata only — it does not retain the parsed node,
        # so this reaches into the package by path. That is a stopgap: the
        # surgical-write design needs charts to hold their XML, and when stage
        # 3 decides how (a `raw` field, or a lookup keyed on `path`), these
        # three lines become one accessor. Don't copy this pattern elsewhere.
        _chart_root(xf, c) = xml_root_element(get_xml_data(xf, c.path))

        XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx")) do xf
            wb = XLSX.get_workbook(xf)
            c = first(getCharts(xf["Data"]))

            root  = _chart_root(xf, c)                          # c:chartSpace
            chart = first_element_with_tag(root, "chart")
            @test chart !== nothing

            @testset "axis txPr" begin
                plotarea = first_element_with_tag(chart, "plotArea")
                @test plotarea !== nothing

                ax = first_element_with_tag(plotarea, "valAx")
                @test ax !== nothing

                t = parse_drawing_text(wb, ax)
                @test t !== nothing
                @test text_content(t) == ""          # formatting only
                @test isempty(text_runs(t))
                @test is_uniform(t)
                @test t.raw !== nothing              # kept for write-back

                # <a:bodyPr rot="-60000000" spcFirstLastPara="1"
                #  vertOverflow="ellipsis" vert="horz" wrap="square"
                #  anchor="ctr" anchorCtr="1"/>
                @test t.body.rotation == -1000.0
                @test t.body.vertical == "horz"
                @test t.body.wrap == "square"
                @test t.body.anchor == "ctr"
                @test t.body.anchorctr === true
                @test t.body.spcfirstlastpara === true
                @test t.body.vertoverflow == "ellipsis"
                @test t.body.horzoverflow === nothing
                @test t.body.autofit === nothing     # absent, not :none
                @test t.body.insetleft === nothing

                # The font is on a:pPr/a:defRPr — this txPr has no runs.
                @test t.paragraphs[1].props.defprops !== nothing
                p = default_run_props(t)
                @test p === t.paragraphs[1].props.defprops

                # <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike"
                #  kern="1200" baseline="0"/>
                @test p.size == 9.0
                @test p.kern == 12.0
                @test p.under == "none"
                @test p.strike == "noStrike"

                # Written as 0, so explicitly false — not absent.
                @test p.bold === false
                @test p.italic === false
                @test p.baseline == 0.0

                # Not written at all.
                @test p.caps === nothing
                @test p.spacing === nothing
                @test p.lang === nothing
                @test p.line === nothing

                # tx1 + lumMod 65000 / lumOff 35000
                @test p.fill !== nothing
                @test p.fill.kind == :solid
                @test p.fill.fgcolor.rgb == "595959"

                # Theme references, kept as written.
                @test p.latin == "+mn-lt"
                @test p.ea == "+mn-ea"
                @test p.cs == "+mn-cs"
                @test resolve_theme_font(wb, p.latin) !== nothing
            end
            @testset "title rich text" begin
                title = first_element_with_tag(chart, "title")
                if title !== nothing
                    tx = first_element_with_tag(title, "tx")
                    if tx !== nothing
                        t = parse_drawing_text(wb, tx; tag="rich")
                        if t !== nothing
                            # `Chart.title` is parsed independently, so this
                            # cross-checks the two paths agree.
                            @test text_content(t) == c.title
                            @test !isempty(text_runs(t))
                        end
                    end
                end
            end
        end

        @testset "theme fonts" begin
            XLSX.openxlsx(joinpath(data_directory, "chart_theme_colors.xlsx")) do xf
                wb = XLSX.get_workbook(xf)

                fonts = get_theme_fonts(wb)
                @test fonts isa Dict{String,String}
                @test haskey(fonts, "+mn-lt")
                @test haskey(fonts, "+mj-lt")

                @test resolve_theme_font(wb, "+mn-lt") == fonts["+mn-lt"]
                @test resolve_theme_font(wb, "Calibri") == "Calibri"
                @test resolve_theme_font(wb, nothing) === nothing
                @test resolve_theme_font(wb, "+xx-lt") === nothing   # undefined ref
            end
        end
    end
end

@testset "DrawingML shape properties" begin

    XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx")) do xf
        wb = XLSX.get_workbook(xf)

        @testset "solid fill and line" begin
            node = _frag("""<c:spPr xmlns:c="$NS_C" $_NSDECL>
                  <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
                  <a:ln w="19050"><a:solidFill><a:srgbClr val="203864"/></a:solidFill>
                    <a:prstDash val="dash"/></a:ln>
                </c:spPr>""")

            sp = parse_drawing_shape_props(wb, node)
            @test sp isa DrawingShapeProps
            @test sp.fill.kind == :solid
            @test sp.fill.fgcolor.rgb == "4472C4"
            @test sp.line.width == 1.5              # points, converted at parse
            @test sp.line.dash == "dash"
            @test sp.line.fill.fgcolor.rgb == "203864"
            @test has_fill(sp)
            @test has_line(sp)
            @test sp.effects === nothing
            @test sp.bwmode === nothing
            @test sp.raw !== nothing
        end

        @testset "noFill is not absence" begin
            # Chart-area shape: transparent background, no border. Both are
            # deliberate, and neither is the same as omitting the element.
            node = _frag("""<c:spPr xmlns:c="$NS_C" $_NSDECL>
                  <a:noFill/>
                  <a:ln><a:noFill/></a:ln>
                </c:spPr>""")

            sp = parse_drawing_shape_props(wb, node)
            @test sp.fill !== nothing               # present...
            @test sp.fill.kind == :none             # ...and explicitly invisible
            @test !has_fill(sp)

            @test sp.line !== nothing
            @test sp.line.fill.kind == :none
            @test !has_line(sp)
        end

        @testset "absence is inheritance" begin
            node = _frag("""<c:spPr xmlns:c="$NS_C" $_NSDECL/>""")

            sp = parse_drawing_shape_props(wb, node)
            @test sp !== nothing                    # the element exists
            @test sp.fill === nothing               # but sets nothing
            @test sp.line === nothing
            @test !has_fill(sp)
            @test !has_line(sp)
        end

        @testset "parse from parent, and no spPr at all" begin
            ser = _frag("""<c:ser xmlns:c="$NS_C" $_NSDECL>
                  <c:idx val="0"/>
                  <c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>
                </c:ser>""")
            sp = parse_drawing_shape_props(wb, ser)
            @test sp !== nothing
            @test sp.fill.fgcolor.rgb == "156082"   # accent1, as pinned elsewhere

            bare = _frag("""<c:ser xmlns:c="$NS_C"><c:idx val="0"/></c:ser>""")
            @test parse_drawing_shape_props(wb, bare) === nothing

            # chains without a guard
            @test parse_drawing_shape_props(wb, nothing) === nothing
        end

        @testset "effects preserved, not modelled" begin
            node = _frag("""<c:spPr xmlns:c="$NS_C" $_NSDECL>
                  <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                  <a:effectLst><a:outerShdw blurRad="50800" dist="38100"/></a:effectLst>
                </c:spPr>""")

            sp = parse_drawing_shape_props(wb, node)
            @test sp.effects !== nothing
            @test localname(sp.effects) == "effectLst"
        end

        @testset "pattern fill keeps DrawingML vocabulary" begin
            node = _frag("""<c:spPr xmlns:c="$NS_C" $_NSDECL>
                  <a:pattFill prst="ltUpDiag">
                    <a:fgClr><a:srgbClr val="000000"/></a:fgClr>
                    <a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>
                  </a:pattFill>
                </c:spPr>""")

            sp = parse_drawing_shape_props(wb, node)
            @test sp.fill.kind == :pattern
            @test sp.fill.preset == "ltUpDiag"      # not "lightUp"
            @test has_fill(sp)
        end
    end

    @testset "fixture: series spPr" begin
        XLSX.openxlsx(joinpath(data_directory, "chart_basic.xlsx")) do xf
            wb = XLSX.get_workbook(xf)
            c = first(getCharts(xf["Data"]))

            root  = xml_root_element(get_xml_data(xf, c.path))
            chart = first_element_with_tag(root, "chart")
            plotarea = first_element_with_tag(chart, "plotArea")
            barchart = first_element_with_tag(plotarea, "barChart")
            @test barchart !== nothing

            ser = first_element_with_tag(barchart, "ser")
            @test ser !== nothing

            sp = parse_drawing_shape_props(wb, ser)
            @test sp !== nothing
            @test sp.raw !== nothing

            @test has_fill(sp)
            @test sp.fill.kind == :solid
            @test sp.fill.fgcolor.rgb == "156082"     # accent1

            # Excel writes an explicit "no border", which is not the same as
            # omitting a:ln. has_line must be false while sp.line is not nothing.
            @test sp.line !== nothing
            @test sp.line.fill.kind == :none
            @test !has_line(sp)

            @test sp.effects !== nothing
            @test localname(sp.effects) == "effectLst"
            if sp !== nothing
                @test sp.raw !== nothing
            end
        end
    end
end
