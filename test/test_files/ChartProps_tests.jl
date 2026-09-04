#
# Tests for the chart appearance accessors in `src/chartprops.jl` — stage 3.
#
# These test node-finding, not DrawingML parsing: that a given accessor reaches
# the right element and hands it to the right parser. DrawingML parsing itself
# is covered in DrawingML_tests.jl, and chart discovery and data extraction in
# Charts_tests.jl. A failure here should mean "the accessor looked in the wrong
# place", not "the colour resolved wrongly".
#
# A few theme-colour values are asserted anyway, because a colour reached
# through an accessor exercises the whole chain, and a regression in either
# layer should fail in both files for different reasons.
#
# ---------------------------------------------------------------------------
# Fixture inventory
#
# chart_appearance.xlsx — built in Excel, one sheet, 3x3 of small integers.
# Every element below is deliberate; if the file is ever regenerated, it needs
# all of this or the tests lose their point.
#
#   Combo chart, two groups:
#     c:barChart   series 1 "2024", series 2 "2025"   axIds 612078287, 460195247
#     c:lineChart  series 3 "2026"                    axIds 1926562432, 1773317264
#
#   Four axes, exercising role filtering, id lookup and the deleted case:
#     catAx 612078287  @b  title "Horizontal"  spPr with noFill + real a:ln
#     valAx 460195247  @l  title "Primary"     majorGridlines with spPr
#     valAx 1773317264 @r  title "Secondary"   majorTickMark="out", crosses="max"
#     catAx 1926562432 @b  delete="1"          no spPr, no txPr, no title
#
#   No axis title is bound to a cell, so getAxisTitleRef is nothing throughout —
#   as is getChartTitleRef. The c:strRef path for titles is untested.
#
#   Series 1: spPr accent1 (a:ln present but noFill — has_line false because the
#     line has no fill, not because the element is absent). Series-level
#     c:dLbls with txPr, sz 1050, accent1 lumMod 75000 -> 104862. One c:dPt at
#     idx 1 filled plain red FF0000. Three c:dLbl —
#       idx 0  retyped "Best", carries BOTH c:tx/c:rich and c:txPr in 00B0F0,
#              and an a:r/a:rPr overriding the paragraph defRPr
#       idx 1  dragged: c:layout/c:manualLayout, no c:dLblPos
#       idx 2  c:delete val="1" and nothing else
#     Plus a linear c:trendline, dispRSqr and dispEq both 1, sysDot dash at
#     1.5pt, with a c:trendlineLbl carrying spPr and txPr but no c:tx.
#
#   Series 2: spPr accent2. Series-level c:dLbls with showVal="1" and a txPr at
#     sz 900, tx1 lumMod 75000 / lumOff 25000 -> 404040 — a different level from
#     series 1, so the two labelled series distinguish rather than duplicate.
#     Fixed-value error bars, c:val val="0.5", no c:errDir (Excel omits it on a
#     bar chart).
#
#   Series 3: line, spPr with a:ln only and no fill element at all — the mirror
#     of series 1. c:marker diamond size 9 with its own spPr. One c:dPt at idx 2
#     whose point-level spPr is a 2.25pt line with no fill (the line segment),
#     and whose c:marker/c:spPr carries fill FF0000 and line 00B050 — so the
#     "fill of a data point" lives on the marker, not the point.
#
#   Chart space: solid bg1 fill and a real line at 0.75pt — the only element in
#   the file with both. Legend @b, txPr sz 900. Chart title with spPr and txPr
#   (sz 1400, tx1 lumMod 65000 / lumOff 35000 -> 595959) but no c:tx, so Excel
#   generates the text. c:chartSpace/c:txPr present but empty, which is the top
#   rung of the inheritance cascade saying nothing.
#
#   Not present, so untested here: c:dropLines, c:hiLowLines, c:upDownBars,
#   c:serLines, c:bandFmts; a c:majorGridlines with no c:spPr (the three-state
#   "on but unformatted" case); a polynomial or moving-average trendline, which
#   would exercise c:order and c:period; custom error bars with c:plus/c:minus.
#
# chart_basic.xlsx        — 2-series bar chart, title "Revenue by Region" as
#                           literal c:rich text, so getChartTitleText and
#                           parse_chart_title can be cross-checked.
# chart_theme_colors.xlsx — 6 series: accent1 (156082) and five variants,
#                           series 5 being accent1 lumMod 75000 -> 104862, the
#                           same value series 1's labels carry in the fixture
#                           above by a different route.
# chart_mixed.xlsx        — a c: chart and a cx: chart on one sheet; the c:
#                           accessors are typed to Chart and must not accept a
#                           ChartEx.
# ---------------------------------------------------------------------------

const _attr                  = XLSX._attr
const first_element_with_tag = XLSX.first_element_with_tag
const elements_with_tag      = XLSX.elements_with_tag
const localname              = XLSX.localname
const XLSXError              = XLSX.XLSXError
const getSeriesFill         = XLSX.getSeriesFill
const getSeriesLine         = XLSX.getSeriesLine
const getMarkerFill  = XLSX.getMarkerFill
const getLabelTextProp     = XLSX.getLabelTextProp
const parse_drawing_fill     = XLSX.parse_drawing_fill
const FormatSite             = XLSX.FormatSite
const Effective              = XLSX.Effective
const has_line               = XLSX.has_line
const getSeriesShapeProps     = XLSX.getSeriesShapeProps
@testset "ChartProps" begin

    f  = XLSX.readxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
    c  = XLSX.getCharts(f)[1]

    @testset "reaching the XML" begin
        root = XLSX.chart_root(c)
        @test localname(root) == "chartSpace"

        # Node identity is the write handle: the accessor must return the node
        # that lives in the tree writexlsx serializes, not a copy. Everything
        # in stage 4 depends on this.
        a = XLSX.getSeriesShapeProps(c, 1)
        b = XLSX.getSeriesShapeProps(c, 1)
        @test a.raw === b.raw
        @test a.raw === first_element_with_tag(c.series[1].raw, "spPr")

        # Two Chart objects from two getCharts calls share their nodes.
        c2 = XLSX.getCharts(f)[1]
        @test XLSX.getSeriesShapeProps(c2, 1).raw === a.raw

        @test_throws XLSX.XLSXError XLSX.getSeriesShapeProps(c, 4)
        @test_throws XLSX.XLSXError XLSX.getSeriesShapeProps(c, 0)
    end

    @testset "series appearance" begin
        @test length(c.series) == 3

        sp1 = XLSX.getSeriesShapeProps(c, 1)
        @test !isnothing(sp1)
        @test sp1.fill.kind == :solid
        @test sp1.fill.fgcolor.rgb == "156082"          # accent1, verified in Excel
        @test XLSX.has_fill(sp1)
        @test !XLSX.has_line(sp1)                        # no a:ln written at all
        @test !isnothing(sp1.line)
        @test sp1.line.fill.kind == :none

        # Series 3 is a line: a:ln present, no fill element. The mirror image.
        sp3 = XLSX.getSeriesShapeProps(c, 3)
        @test XLSX.has_line(sp3)
        @test !XLSX.has_fill(sp3)
        @test isnothing(sp3.fill)
    end

    @testset "series data labels" begin
        tx1 = XLSX.getSeriesLabelTextProps(c, 1)
        @test !isnothing(tx1)
        rp = XLSX.default_run_props(tx1)
        @test rp.size == 10.5                            # sz="1050", hundredths -> points
        @test rp.fill.fgcolor.rgb == "104862"            # accent1 + lumMod 75000

        # Series 2 and 3 have no series-level dLbls.
        tx2 = XLSX.getSeriesLabelTextProps(c, 2)
        @test !isnothing(tx2)
        rp2 = XLSX.default_run_props(tx2)
        @test rp2.size == 9.0
        @test rp2.fill.fgcolor.rgb == "404040"   # tx1 + lumMod 75000 / lumOff 25000
        @test isnothing(XLSX.getSeriesLabelTextProps(c, 3))
    end

    @testset "series markers" begin
        @test isnothing(XLSX.getSeriesMarker(c, 1))        # bar series
        @test isnothing(XLSX.getSeriesMarker(c, 2))

        mk = XLSX.getSeriesMarker(c, 3)
        @test mk.symbol == :diamond
        @test mk.size == 9
        @test !isnothing(mk.shape)
        @test XLSX.has_fill(mk.shape)
    end

    @testset "chart groups" begin
        gs = XLSX.getChartGroups(c)
        @test length(gs) == 2
        @test [g.kind for g in gs] == [:barChart, :lineChart]
        @test gs[1].axids == [612078287, 460195247]
        @test gs[2].axids == [1926562432, 1773317264]

        # Group-level dLbls carry only the show* flags here, no txPr.
        @test isnothing(XLSX.getGroupLabelTextProps(c, gs[1]))
        @test isnothing(XLSX.getGroupLabelTextProps(c, gs[2]))

        @test [a.kind for a in XLSX.getGroupAxes(c, gs[1])] == [:catAx, :valAx]
        @test [a.axid for a in XLSX.getGroupAxes(c, gs[2])] == [1926562432, 1773317264]

        # The group is what ties a series to its axes.
        @test XLSX.getSeriesGroup(c, 1).kind == :barChart
        @test XLSX.getSeriesGroup(c, 3).kind == :lineChart
        @test [a.pos for a in XLSX.getSeriesAxes(c, 1)] == [:b, :l]
        @test [a.pos for a in XLSX.getSeriesAxes(c, 3)] == [:b, :r]
    end

    @testset "axes: identity and lookup" begin
        axes = XLSX.getChartAxes(c)
        @test length(axes) == 4
        @test [a.kind for a in axes] == [:catAx, :valAx, :valAx, :catAx]
        @test [a.axid for a in axes] == [612078287, 460195247, 1773317264, 1926562432]
        @test [a.pos for a in axes] == [:b, :l, :r, :b]

        # A combo chart has two value axes, which is why this returns a vector.
        @test length(XLSX.getChartAxes(c, :value)) == 2
        @test length(XLSX.getChartAxes(c, :category)) == 2
        @test isempty(XLSX.getChartAxes(c, :date))
        @test_throws XLSX.XLSXError XLSX.getChartAxes(c, :nonsense)

        @test XLSX.getChartAxis(c, 1773317264).pos == :r
        @test_throws XLSX.XLSXError XLSX.getChartAxis(c, 1)

        # crossAx forms two closed pairs.
        @test XLSX.getAxisPartner(c, axes[1]).axid == 460195247
        @test XLSX.getAxisPartner(c, axes[2]).axid == 612078287
        @test XLSX.getAxisPartner(c, axes[3]).axid == 1926562432
        @test XLSX.getAxisPartner(c, axes[4]).axid == 1773317264
    end

    @testset "axes: deleted" begin
        del = XLSX.getChartAxis(c, 1926562432)
        @test del.deleted
        @test all(!a.deleted for a in XLSX.getChartAxes(c) if a.axid != 1926562432)

        # Excel strips formatting from a deleted axis but keeps its structure.
        @test isnothing(XLSX.getAxisShapeProps(c, del))
        @test isnothing(XLSX.getAxisTextProps(c, del))
        @test isnothing(XLSX.getAxisTitleText(c, del))
        @test isnothing(XLSX.getAxisGridlines(c, del))
        @test XLSX.getAxisLabelOffset(c, del) == 100      # scalars survive
        @test !isnothing(XLSX.getAxisPartner(c, del))
    end

    @testset "axes: appearance" begin
        cat, pri, sec = XLSX.getChartAxis(c, 612078287),
                        XLSX.getChartAxis(c, 460195247),
                        XLSX.getChartAxis(c, 1773317264)

        @test XLSX.text_content(XLSX.getAxisTitleText(c, cat)) == "Horizontal"
        @test XLSX.text_content(XLSX.getAxisTitleText(c, pri)) == "Primary"
        @test XLSX.text_content(XLSX.getAxisTitleText(c, sec)) == "Secondary"

        # No fixture has a title bound to a cell.
        @test isnothing(XLSX.getAxisTitleRef(c, cat))

        # noFill plus a real line, versus noFill on both: same has_* answers
        # from different XML, and both explicit rather than absent.
        spc = XLSX.getAxisShapeProps(c, cat)
        @test !XLSX.has_fill(spc) && XLSX.has_line(spc)
        @test spc.fill.kind == :none                     # explicit <a:noFill/>

        spp = XLSX.getAxisShapeProps(c, pri)
        @test !XLSX.has_fill(spp) && !XLSX.has_line(spp)

        @test !isnothing(XLSX.getAxisGridlines(c, pri))
        @test isnothing(XLSX.getAxisGridlines(c, cat))
        @test isnothing(XLSX.getAxisGridlines(c, pri; minor=true))
    end

    @testset "axes: scalars" begin
        cat, pri, sec = XLSX.getChartAxis(c, 612078287),
                        XLSX.getChartAxis(c, 460195247),
                        XLSX.getChartAxis(c, 1773317264)

        @test XLSX.getAxisNumberFormatCode(c, cat) == "General"
        @test XLSX.getAxisNumberFormatLinked(c, cat) === true

        @test XLSX.getAxisMajorTickMark(c, cat) == :none
        @test XLSX.getAxisMajorTickMark(c, sec) == :out
        @test XLSX.getAxisMinorTickMark(c, cat) == :none
        @test XLSX.getAxisTickLabelPos(c, cat) == :nextTo

        @test XLSX.getAxisOrientation(c, cat) == :minMax
        @test isnothing(XLSX.getAxisMin(c, pri))           # automatic scaling
        @test isnothing(XLSX.getAxisMax(c, pri))
        @test isnothing(XLSX.getAxisLogBase(c, pri))

        @test XLSX.getAxisCrosses(c, pri) == :autoZero
        @test XLSX.getAxisCrosses(c, sec) == :max
        @test isnothing(XLSX.getAxisCrossesAt(c, pri))    # mutually exclusive

        @test isnothing(XLSX.getAxisMajorUnit(c, pri))
        @test isnothing(XLSX.getAxisMinorUnit(c, pri))

        @test XLSX.getAxisLabelAlign(c, cat) == :ctr
        @test XLSX.getAxisLabelOffset(c, cat) == 100
        @test XLSX.getAxisMultiLevelLabels(c, cat) === true   # noMultiLvlLbl="0"
        @test XLSX.getAxisCrossBetween(c, pri) == :between

        # Kind-specific accessors throw rather than returning nothing, so that
        # nothing keeps meaning "not written".
        @test_throws XLSX.XLSXError XLSX.getAxisCrossBetween(c, cat)
        @test_throws XLSX.XLSXError XLSX.getAxisLabelOffset(c, pri)
        @test_throws XLSX.XLSXError XLSX.getAxisLabelAlign(c, pri)
        @test_throws XLSX.XLSXError XLSX.getAxisMajorUnit(c, cat)
    end

    @testset "chart space level" begin
        # Title: txPr present, no c:tx at all — Excel generates the text.
        @test !isnothing(XLSX.getChartTitleNode(c))
        @test isnothing(XLSX.getChartTitleText(c))
        @test isnothing(XLSX.getChartTitleRef(c))
        @test XLSX.getAutoTitleDeleted(c) === false
        ttp = XLSX.getChartTitleTextProps(c)
        @test XLSX.default_run_props(ttp).size == 14.0
        @test XLSX.default_run_props(ttp).fill.fgcolor.rgb == "595959"  # tx1 +lumMod/lumOff

        @test XLSX.getLegendPos(c) == :b
        @test XLSX.getLegendOverlay(c) === false
        @test XLSX.default_run_props(XLSX.getLegendTextProps(c)).size == 9.0

        pa = XLSX.getPlotAreaShapeProps(c)
        @test !XLSX.has_fill(pa) && !XLSX.has_line(pa)

        # The only element in the file with both a real fill and a real line.
        cs = XLSX.getChartSpaceShapeProps(c)
        @test XLSX.has_fill(cs) && XLSX.has_line(cs)
        @test cs.fill.fgcolor.rgb == "FFFFFF"            # bg1 -> lt1
        @test cs.line.width == 0.75                      # w="9525" EMU -> points

        # The top of the cascade: present, but says nothing.
        cst = XLSX.getChartSpaceTextProps(c)
        @test !isnothing(cst)
        @test isnothing(XLSX.default_run_props(cst).size)
    end

    @testset "data points" begin
        @test length(XLSX.getSeriesDataPoints(c, 1)) == 1
        @test isempty(XLSX.getSeriesDataPoints(c, 2))
        @test length(XLSX.getSeriesDataPoints(c, 3)) == 1

        d1 = XLSX.getSeriesDataPoint(c, 1, 2)             # 1-based position
        @test !isnothing(d1)
        @test d1.idx == 1                                # c:idx is 0-based
        @test d1.invert_if_negative === false            # written by Excel
        @test XLSX.getDataPointShapeProps(c, d1).fill.fgcolor.rgb == "FF0000"
        @test isnothing(XLSX.getDataPointMarker(c, d1))

        @test isnothing(XLSX.getSeriesDataPoint(c, 1, 1))
        @test isnothing(XLSX.getSeriesDataPoint(c, 1, 3))
        @test_throws XLSX.XLSXError XLSX.getSeriesDataPoint(c, 1, 0)

        # A point that overrides only its marker: the point-level spPr is the
        # line segment, which has no fill; the colour lives on the marker.
        d3 = XLSX.getSeriesDataPoint(c, 3, 3)
        @test d3.idx == 2
        @test isnothing(d3.invert_if_negative)           # meaningless on a line
        sp = XLSX.getDataPointShapeProps(c, d3)
        @test isnothing(sp.fill)                         # absent, not noFill
        @test sp.line.width == 2.25
        mk = XLSX.getDataPointMarker(c, d3)
        @test mk.symbol == :diamond && mk.size == 9
        @test mk.shape.fill.fgcolor.rgb == "FF0000"

        # idx round-trips through the 1-based lookup to the same node.
        for i in 1:length(c.series), d in XLSX.getSeriesDataPoints(c, i)
            @test XLSX.getSeriesDataPoint(c, i, d.idx + 1).raw === d.raw
        end
    end

    @testset "individual data labels" begin
        dls = XLSX.getSeriesDataLabels(c, 1)
        @test length(dls) == 3
        @test [d.idx for d in dls] == [0, 1, 2]
        @test isempty(XLSX.getSeriesDataLabels(c, 2))

        # Retyped label: literal text, and formatting in both c:rich and c:txPr.
        r = XLSX.getSeriesDataLabel(c, 1, 1)
        @test XLSX.text_content(XLSX.getDataLabelText(c, r)) == "Best"
        @test !isnothing(XLSX.getDataLabelTextProps(c, r))
        @test XLSX.default_run_props(XLSX.getDataLabelText(c, r)).fill.fgcolor.rgb == "00B0F0"
        @test isnothing(XLSX.getDataLabelOffset(c, r))
        @test isnothing(XLSX.getDataLabelPosition(c, r))

        # Dragged label: a manual layout offset, no dLblPos.
        m = XLSX.getSeriesDataLabel(c, 1, 2)
        off = XLSX.getDataLabelOffset(c, m)
        @test off.x ≈ -0.038888888888888994
        @test off.y ≈ -0.06481481481481485
        @test isnothing(XLSX.getDataLabelPosition(c, m))
        @test isnothing(XLSX.getDataLabelText(c, m))

        # Deleted label: c:delete and nothing else.
        x = XLSX.getSeriesDataLabel(c, 1, 3)
        @test x.delete === true
        @test isnothing(XLSX.getDataLabelText(c, x))
        @test isnothing(XLSX.getDataLabelTextProps(c, x))

        # Absent delete means shown, and is distinct from an explicit false.
        @test r.delete === nothing
        @test m.delete === nothing
        @test length(filter(d -> d.delete !== true, dls)) == 2
    end

    @testset "trendlines" begin
        @test isempty(XLSX.getSeriesTrendlines(c, 2))
        ts = XLSX.getSeriesTrendlines(c, 1)
        @test length(ts) == 1

        t = ts[1]
        @test t.kind == :linear
        @test t.disp_rsqr === true
        @test t.disp_eq === true
        # A plain linear fit writes almost nothing else.
        @test isnothing(t.name)
        @test isnothing(t.order)
        @test isnothing(t.period)
        @test isnothing(t.forward)
        @test isnothing(t.backward)
        @test isnothing(t.intercept)

        sp = XLSX.getTrendlineShapeProps(c, t)
        @test sp.line.width == 1.5                       # w="19050"
        @test sp.line.dash == "sysDot"                   # element, not attribute
        @test isnothing(sp.fill)

        # The label carries formatting but no typed text.
        @test isnothing(XLSX.getTrendlineLabelText(c, t))
        @test !isnothing(XLSX.getTrendlineLabelTextProps(c, t))
        @test !isnothing(XLSX.getTrendlineLabelShapeProps(c, t))
    end

    @testset "error bars" begin
        @test isempty(XLSX.getSeriesErrorBars(c, 1))
        es = XLSX.getSeriesErrorBars(c, 2)
        @test length(es) == 1

        e = es[1]
        @test isnothing(e.direction)                     # omitted on a bar chart
        @test e.bar_type == :both
        @test e.value_type == :fixedVal
        @test e.value == 0.5
        @test e.no_end_cap === false

        sp = XLSX.getErrorBarsShapeProps(c, e)
        @test XLSX.has_line(sp)
        @test sp.fill.kind == :none                      # explicit noFill

        # Only cust bars carry plus/minus references.
        @test XLSX.getErrorBarsCustomRefs(c, e) == (plus = nothing, minus = nothing)
    end

    @testset "other fixtures" begin
        fb = XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))
        cb = XLSX.getCharts(fb)[1]

        # A literal title, reached two ways: the accessor and parse_chart_title.
        @test XLSX.text_content(XLSX.getChartTitleText(cb)) == "Revenue by Region"
        @test cb.title == "Revenue by Region"
        @test length(XLSX.getChartGroups(cb)) == 1
        @test isempty(XLSX.getSeriesDataPoints(cb, 1))

        # Six theme variants, all reached through the accessor.
        ft = XLSX.readxlsx(joinpath(data_directory, "chart_theme_colors.xlsx"))
        ct = XLSX.getCharts(ft)[1]
        @test length(ct.series) == 6
        @test XLSX.getSeriesShapeProps(ct, 1).fill.fgcolor.rgb == "156082"
        @test XLSX.getSeriesShapeProps(ct, 5).fill.fgcolor.rgb == "104862"

        # A cx: chart is a ChartEx, and the c: accessors are not defined for it.
        fm = XLSX.readxlsx(joinpath(data_directory, "chart_mixed.xlsx"))
        cms = XLSX.getCharts(fm)
        @test any(x -> x isa XLSX.ChartEx, cms)
        cx = first(x for x in cms if x isa XLSX.ChartEx)
        @test_throws MethodError XLSX.getSeriesShapeProps(cx, 1)
    end

    @testset "fill cascade" begin
        # Series 3: spPr with a:ln only, no fill element — must not stop the walk.
        e = getSeriesFill(c, 3)
        @test isnothing(e.value) && isnothing(e.site)
        @test length(e.chain) == 1 && !isnothing(e.chain[1].props)   # rung exists, fill doesn't

        # Series 1: solid accent1 on the series.
        e = getSeriesFill(c, 1)
        @test e.site.level === :series && e.value.kind === :solid

        # Series 1 point idx 1 (position 2): plain red, overrides the series.
        e = getSeriesFill(c, 1, point=2)
        @test e.site.level === :point && e.value.fgcolor.rgb == "FF0000"

        # Series 3 point idx 2 (position 3): point spPr is the line segment, no
        # fill — the fill Excel shows is on the marker.
        e = getSeriesFill(c, 3, point=3)
        @test isnothing(e.value)
        @test length(e.chain) == 2 && all(s -> s.kind === :shape, e.chain)

        m = getMarkerFill(c, 3, 3)
        @test m.site.level === :point && m.site.kind === :marker && m.value.fgcolor.rgb == "FF0000"

        # Series 3's own marker (diamond, size 9) has its own spPr, so it is the
        # second rung and answers when the point has none.
        @test length(m.chain) == 2
    end

    @testset "line cascade" begin
        # Series 1: a:ln present but noFill — has_line false. The line element IS
        # written, so the walk stops here; the resolved DrawingLine has a :none fill.
        e = getSeriesLine(c, 1)
        @test e.site.level === :series
        @test !isnothing(e.value) && e.value.fill.kind === :none
        @test !has_line(getSeriesShapeProps(c, 1))
    end
    @testset "empty txPr does not answer" begin
        e = getLabelTextProp(c, 3, :size)          # series 3 has no dLbls at all
        @test isnothing(e.value) && isnothing(e.site)
        @test length(e.chain) == 4                    # series, group, plotarea, chartspace
        @test e.chain[1].level === :series && isnothing(e.chain[1].props)
        @test e.chain[end].level === :chartspace && !isnothing(e.chain[end].props)  # present, empty
    end

    @testset "text cascade resolves at the series" begin
        e = getLabelTextProp(c, 1, :size)
        @test e.site.level === :series && e.value == 10.5
        e = getLabelTextProp(c, 2, :size)
        @test e.site.level === :series && e.value == 9.0
    end
end

