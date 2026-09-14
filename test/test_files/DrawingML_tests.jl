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
## =============================================================================

const XLSXError = XLSX.XLSXError
const localname = XLSX.localname

const _attr           = XLSX._attr
const get_attr        = XLSX.get_attr
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
const parse_drawing_color = XLSX.parse_drawing_color
const default_run_props   = XLSX.default_run_props
const first_run_props     = XLSX.first_run_props
const is_uniform          = XLSX.is_uniform
const text_content        = XLSX.text_content
const text_runs           = XLSX.text_runs
const DrawingText         = XLSX.DrawingText
const get_theme_fonts     = XLSX.get_theme_fonts
const resolve_theme_font  = XLSX.resolve_theme_font
const resolve_color_base  = XLSX.resolve_color_base
const xml_root_element    = XLSX.xml_root_element
const get_xml_data        = XLSX.get_xml_data
const getCharts           = XLSX.getCharts

const parse_drawing_shape_props = XLSX.parse_drawing_shape_props
const DrawingShapeProps         = XLSX.DrawingShapeProps
const has_fill                  = XLSX.has_fill
const has_line                  = XLSX.has_line

const SchemeColor = XLSX.SchemeColor
const _scheme_color_node = XLSX._scheme_color_node
const _solid_fill_node = XLSX._solid_fill_node
const _srgb_color_node = XLSX._srgb_color_node

const apply_drawingml_transforms = XLSX.apply_drawingml_transforms

const _ln_with_width = XLSX._ln_with_width
const _ln_with_dash = XLSX._ln_with_dash
const _ln_with_cap = XLSX._ln_with_cap
const _ln_with_compound = XLSX._ln_with_compound
const _ln_with_join = XLSX._ln_with_join
const _ln_with_miter_limit = XLSX._ln_with_miter_limit
const _ln_with_color = XLSX._ln_with_color

const _text_from         = XLSX._text_from
const _fill_node_from    = XLSX._fill_node_from
const _color_node_from   = XLSX._color_node_from
const DrawingText        = XLSX.DrawingText
const DrawingParagraph   = XLSX.DrawingParagraph
const DrawingRun         = XLSX.DrawingRun
const DrawingRunProps    = XLSX.DrawingRunProps
const DrawingParaProps   = XLSX.DrawingParaProps
const DrawingBodyProps   = XLSX.DrawingBodyProps
const DrawingFill        = XLSX.DrawingFill
const DrawingLine        = XLSX.DrawingLine
const DrawingColor       = XLSX.DrawingColor

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

@testset "SchemeColor" begin
    @testset "construction" begin
        @test SchemeColor(:accent1).token === :accent1
        @test isempty(SchemeColor(:accent1).transforms)
        @test SchemeColor(:accent1; lumMod = 75).transforms == [:lumMod => 75.0]
        @test SchemeColor(:tx1; lumMod = 65, lumOff = 35).transforms ==
              [:lumMod => 65.0, :lumOff => 35.0]
        # integers convert
        @test SchemeColor(:accent1; alpha = 80).transforms == [:alpha => 80.0]
        # vector form takes any order
        @test SchemeColor(:accent2, [:shade => 50.0, :alpha => 80.0]).transforms ==
              [:shade => 50.0, :alpha => 80.0]
    end

    @testset "validation" begin
        @test_throws XLSXError SchemeColor(:accent7)
        @test_throws XLSXError SchemeColor(:acccent1)          # typo, not silently accepted
        @test_throws XLSXError SchemeColor(:accent1, [:lumMud => 75.0])
        @test_throws XLSXError SchemeColor(:theme1)            # spreadsheet vocabulary, not DrawingML
    end

    @testset "aliases" begin
        # preserved as written
        @test SchemeColor(:lt1).token === :lt1
        @test SchemeColor(:dk2).token === :dk2
        # but equal and equally hashed
        @test SchemeColor(:lt1) == SchemeColor(:bg1)
        @test SchemeColor(:dk1) == SchemeColor(:tx1)
        @test SchemeColor(:lt2) == SchemeColor(:bg2)
        @test SchemeColor(:dk2) == SchemeColor(:tx2)
        @test hash(SchemeColor(:lt1)) == hash(SchemeColor(:bg1))
        @test isequal(SchemeColor(:lt1), SchemeColor(:bg1))
        # so one Dict slot, not two
        d = Dict(SchemeColor(:lt1) => 1); d[SchemeColor(:bg1)] = 2
        @test length(d) == 1
        # aliases with matching transforms are still equal
        @test SchemeColor(:lt1; lumMod = 50) == SchemeColor(:bg1; lumMod = 50)
        # different slots are not
        @test SchemeColor(:accent1) != SchemeColor(:accent2)
    end

    @testset "transform order is significant" begin
        a = SchemeColor(:accent1; lumMod = 75, lumOff = 25)
        b = SchemeColor(:accent1, [:lumOff => 25.0, :lumMod => 75.0])
        @test a.transforms != b.transforms
        @test a != b
        @test SchemeColor(:accent1; lumMod = 75) != SchemeColor(:accent1; lumMod = 50)
        @test SchemeColor(:accent1) != SchemeColor(:accent1; lumMod = 75)
    end

    xf = XLSX.openxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
    wb = XLSX.get_workbook(xf[1])
    @testset "SchemeColor round trip" begin

        # resolve_color_base does the theme lookup only; the transforms are applied
        # separately, so composing the two is what a reader effectively does.
        resolved(wb, n) = first(apply_drawingml_transforms(
            resolve_color_base(wb, n),
            parse_drawing_color(wb, n).transforms))

        # accent1, untransformed
        n  = _scheme_color_node(SchemeColor(:accent1))
        @test XML.write(n) == """<a:schemeClr val="accent1"/>"""
        dc = parse_drawing_color(wb, n)
        @test dc.kind === :scheme && dc.val == "accent1"

        # accent1 + lumMod 75% — 104862 against the Office theme, verified in stage 2
        n = _scheme_color_node(SchemeColor(:accent1; lumMod = 75))
        @test occursin("""val="75000\"""", XML.write(n))
        dc = parse_drawing_color(wb, n)
        @test length(dc.transforms) == 1
        @test resolve_color_base(wb, n) == "156082"

        # tx1 + lumMod 65 / lumOff 35 — 595959, and order is preserved through the XML
        n = _scheme_color_node(SchemeColor(:tx1; lumMod = 65, lumOff = 35))
        @test localname.(collect(XML.eachelement(n))) == ["lumMod", "lumOff"]
        @test resolve_color_base(wb, n) == "000000"

        # what we build parses back to what we built
        for sc in (SchemeColor(:accent1), SchemeColor(:accent1; lumMod = 75),
                SchemeColor(:tx1; lumMod = 65, lumOff = 35),
                SchemeColor(:accent2, [:shade => 50.0, :alpha => 80.0]))
            dc = parse_drawing_color(wb, _scheme_color_node(sc))
            @test dc.val == String(sc.token)
            @test length(dc.transforms) == length(sc.transforms)
        end

        @test resolve_color_base(wb, _scheme_color_node(SchemeColor(:accent1))) == "156082"
        @test resolved(wb, _scheme_color_node(SchemeColor(:accent1; lumMod = 75))) == "104862"
        @test resolved(wb, _scheme_color_node(SchemeColor(:tx1; lumMod = 65, lumOff = 35))) == "595959"

        # solidFill wrapper
        f = parse_drawing_fill(wb, _solid_fill_node(_scheme_color_node(SchemeColor(:accent1))))
        @test f.kind === :solid && f.fgcolor.val == "accent1"

        # alpha rides alongside the hex rather than altering it
        n = _scheme_color_node(SchemeColor(:accent1; alpha = 50))
        hex, alpha = apply_drawingml_transforms(
            resolve_color_base(wb, n), parse_drawing_color(wb, n).transforms)
        @test hex == "156082"
        @test alpha ≈ 0.5
    end

    @testset "srgb color nodes" begin
        # 8-digit hex passes through, split into six digits plus alpha
        @test XML.write(_srgb_color_node("FFFF0000")) == """<a:srgbClr val="FF0000"/>"""
        @test XML.write(_srgb_color_node("ffff0000")) == """<a:srgbClr val="FF0000"/>"""   # case

        # Colors.jl names
        @test XML.write(_srgb_color_node("red"))   == """<a:srgbClr val="FF0000"/>"""
        @test XML.write(_srgb_color_node(:red))    == """<a:srgbClr val="FF0000"/>"""
        @test XML.write(_srgb_color_node("grey"))  == XML.write(_srgb_color_node("gray"))

        # opaque emits no alpha child
        @test isnothing(_srgb_color_node("FF0000FF").children) ||
            isempty(_srgb_color_node("FF0000FF").children)

        # partial alpha becomes a transform
        n = _srgb_color_node("800000FF")
        @test localname.(collect(XML.eachelement(n))) == ["alpha"]
        @test get_attr(first(XML.eachelement(n)), "val") == "50196"   # 128/255

        # colorants, opaque and transparent
        @test XML.write(_srgb_color_node(Colors.RGB(1, 0, 0))) == """<a:srgbClr val="FF0000"/>"""
        n = _srgb_color_node(Colors.ARGB(1, 0, 0, 0.5))
        @test get_attr(n, "val") == "FF0000"
        @test !isempty(XML.children(n))

        # invalid names throw get_color's message
        @test_throws XLSXError _srgb_color_node("notacolor")

        # round trip through the parser
        dc = parse_drawing_color(wb, _srgb_color_node("red"))
        @test dc.kind === :srgb && dc.rgb == "FF0000"
    end

    @testset "_ln_with_*" begin
        pfx = Dict(NS_A => "a")
        ln() = XML.Element("a:ln")

        @test get_attr(_ln_with_width(ln(), 2), "w") == "25400"
        @test_throws XLSXError _ln_with_width(ln(), 2000)
        @test_throws XLSXError _ln_with_width(ln(), -1)
        @test get_attr(_ln_with_width(_ln_with_width(ln(), 2), :inherit), "w") == ""

        # Excel names and DrawingML names both work, and resolve to the same thing
        @test XML.write(_ln_with_dash(ln(), :roundDot, pfx)) ==
            XML.write(_ln_with_dash(ln(), :sysDot, pfx))
        @test_throws XLSXError _ln_with_dash(ln(), :dotted, pfx)

        @test get_attr(_ln_with_cap(ln(), :round), "cap") == "rnd"
        @test get_attr(_ln_with_cap(ln(), :rnd), "cap") == "rnd"
        @test get_attr(_ln_with_compound(ln(), :double), "cmpd") == "dbl"

        # join, and the limit that depends on it
        l = _ln_with_join(ln(), :miter, pfx)
        @test localname.(collect(XML.eachelement(l))) == ["miter"]
        @test get_attr(first(XML.eachelement(_ln_with_miter_limit(l, 8))), "lim") == "800000"
        @test_throws XLSXError _ln_with_miter_limit(ln(), 8)
        @test_throws XLSXError _ln_with_miter_limit(_ln_with_join(ln(), :bevel, pfx), 8)

        # schema order: fill, dash, join
        full = _ln_with_join(_ln_with_dash(_ln_with_color(ln(), "red", pfx), :dash, pfx), :round, pfx)
        @test localname.(collect(XML.eachelement(full))) == ["solidFill", "prstDash", "round"]

        # each choice group replaces rather than accumulates
        @test localname.(collect(XML.eachelement(
            _ln_with_color(_ln_with_color(ln(), "red", pfx), :none, pfx)))) == ["noFill"]
    end

    @testset "text body round trip" begin
        pfx = Dict(NS_A => "a", NS_C => "c")

        rt(t) = parse_drawing_text(wb, _text_from(t, "txPr", pfx))

        # plain text
        t = DrawingText("Revenue by Region")
        g = rt(t)
        @test length(g.paragraphs) == 1
        @test text_content(g) == "Revenue by Region"

        # run properties survive, in both directions
        t = DrawingText(DrawingParagraph(
                DrawingRun("Revenue", props = DrawingRunProps(size = 14.0, bold = true,
                                                            latin = "Calibri"))))
        g = rt(t)
        rp = first_run_props(g)
        @test rp.size ≈ 14.0 && rp.bold === true && rp.latin == "Calibri"
        @test isnothing(rp.italic)          # absent stays absent

        # paragraph defaults
        t = DrawingText(DrawingParagraph("x",
                props = DrawingParaProps(align = "ctr",
                                        defprops = DrawingRunProps(size = 10.5))))
        g = rt(t)
        @test g.paragraphs[1].props.align == "ctr"
        @test g.paragraphs[1].props.defprops.size ≈ 10.5

        # a solid fill on a run, scheme and srgb
        for color in (DrawingColor(:srgb, "FF0000", Pair{Symbol,Int}[], "FF0000", 1.0),
                    DrawingColor(:scheme, "accent1", [:lumMod => 75000], "104862", 1.0))
            t = DrawingText(DrawingParagraph(DrawingRun("x",
                    props = DrawingRunProps(fill = DrawingFill(:solid; fgcolor = color)))))
            f = first_run_props(rt(t)).fill
            @test f.kind === :solid
            @test f.fgcolor.val == color.val
            @test f.fgcolor.transforms == color.transforms
        end

        # a text outline
        t = DrawingText(DrawingParagraph(DrawingRun("x",
                props = DrawingRunProps(line = DrawingLine(width = 1.5, dash = "sysDot")))))
        l = first_run_props(rt(t)).line
        @test l.width ≈ 1.5 && l.dash == "sysDot"

        # body properties, including the three-way autofit
        for (af, extra) in ((:none, ()), (:shape, ()),
                            (:normal, (fontscale = 0.9, linespacereduction = 0.1)))
            t = DrawingText("x"; body = DrawingBodyProps(; rotation = -45.0, anchor = "ctr",
                                                        autofit = af, extra...))
            b = rt(t).body
            @test b.rotation ≈ -45.0 && b.anchor == "ctr" && b.autofit === af
        end
        b = rt(DrawingText("x"; body = DrawingBodyProps(autofit = :normal, fontscale = 0.9))).body
        @test b.fontscale ≈ 0.9

        # several runs keep document order
        t = DrawingText(DrawingParagraph("one", DrawingRun("\n", kind = :br), "two"))
        g = rt(t)
        @test [r.kind for r in g.paragraphs[1].runs] == [:run, :br, :run]
        @test text_content(g) == "one\ntwo"

        # schema order inside a:p and a:defRPr
        n = _text_from(DrawingText(DrawingParagraph("x",
                props = DrawingParaProps(defprops = DrawingRunProps(size = 10.0)))), "txPr", pfx)
        p = first_element_with_tag(n, "p")
        @test localname.(collect(XML.eachelement(p))) == ["pPr", "r"]

        # what cannot be written says so
        @test_throws XLSXError _fill_node_from(DrawingFill(:gradient), pfx)
        @test_throws XLSXError _color_node_from(
            DrawingColor(:scrgb, "", Pair{Symbol,Int}[], "000000", 1.0), pfx)
    end
end