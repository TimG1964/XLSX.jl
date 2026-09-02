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
#   No axis title is bound to a cell, so axis_title_ref is nothing throughout —
#   as is chart_title_ref. The c:strRef path for titles is untested.
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
#                           literal c:rich text, so chart_title_text and
#                           parse_chart_title can be cross-checked.
# chart_theme_colors.xlsx — 6 series: accent1 (156082) and five variants,
#                           series 5 being accent1 lumMod 75000 -> 104862, the
#                           same value series 1's labels carry in the fixture
#                           above by a different route.
# chart_mixed.xlsx        — a c: chart and a cx: chart on one sheet; the c:
#                           accessors are typed to Chart and must not accept a
#                           ChartEx.
# ---------------------------------------------------------------------------

const _attr = XLSX._attr
const first_element_with_tag = XLSX.first_element_with_tag
const elements_with_tag = XLSX.elements_with_tag
const localname = XLSX.localname

@testset "ChartProps" begin

    f  = XLSX.readxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
    c  = XLSX.getCharts(f)[1]

    @testset "reaching the XML" begin
        root = XLSX.chart_root(c)
        @test localname(root) == "chartSpace"

        # Node identity is the write handle: the accessor must return the node
        # that lives in the tree writexlsx serializes, not a copy. Everything
        # in stage 4 depends on this.
        a = XLSX.series_shape_props(c, 1)
        b = XLSX.series_shape_props(c, 1)
        @test a.raw === b.raw
        @test a.raw === first_element_with_tag(c.series[1].raw, "spPr")

        # Two Chart objects from two getCharts calls share their nodes.
        c2 = XLSX.getCharts(f)[1]
        @test XLSX.series_shape_props(c2, 1).raw === a.raw

        @test_throws XLSX.XLSXError XLSX.series_shape_props(c, 4)
        @test_throws XLSX.XLSXError XLSX.series_shape_props(c, 0)
    end

    @testset "series appearance" begin
        @test length(c.series) == 3

        sp1 = XLSX.series_shape_props(c, 1)
        @test !isnothing(sp1)
        @test sp1.fill.kind == :solid
        @test sp1.fill.fgcolor.rgb == "156082"          # accent1, verified in Excel
        @test XLSX.has_fill(sp1)
        @test !XLSX.has_line(sp1)                        # no a:ln written at all
        @test !isnothing(sp1.line)
        @test sp1.line.fill.kind == :none

        # Series 3 is a line: a:ln present, no fill element. The mirror image.
        sp3 = XLSX.series_shape_props(c, 3)
        @test XLSX.has_line(sp3)
        @test !XLSX.has_fill(sp3)
        @test isnothing(sp3.fill)
    end

    @testset "series data labels" begin
        tx1 = XLSX.series_label_text_props(c, 1)
        @test !isnothing(tx1)
        rp = XLSX.default_run_props(tx1)
        @test rp.size == 10.5                            # sz="1050", hundredths -> points
        @test rp.fill.fgcolor.rgb == "104862"            # accent1 + lumMod 75000

        # Series 2 and 3 have no series-level dLbls.
        tx2 = XLSX.series_label_text_props(c, 2)
        @test !isnothing(tx2)
        rp2 = XLSX.default_run_props(tx2)
        @test rp2.size == 9.0
        @test rp2.fill.fgcolor.rgb == "404040"   # tx1 + lumMod 75000 / lumOff 25000
        @test isnothing(XLSX.series_label_text_props(c, 3))
    end

    @testset "series markers" begin
        @test isnothing(XLSX.series_marker(c, 1))        # bar series
        @test isnothing(XLSX.series_marker(c, 2))

        mk = XLSX.series_marker(c, 3)
        @test mk.symbol == :diamond
        @test mk.size == 9
        @test !isnothing(mk.shape)
        @test XLSX.has_fill(mk.shape)
    end

    @testset "chart groups" begin
        gs = XLSX.chart_groups(c)
        @test length(gs) == 2
        @test [g.kind for g in gs] == [:barChart, :lineChart]
        @test gs[1].axids == [612078287, 460195247]
        @test gs[2].axids == [1926562432, 1773317264]

        # Group-level dLbls carry only the show* flags here, no txPr.
        @test isnothing(XLSX.group_label_text_props(c, gs[1]))
        @test isnothing(XLSX.group_label_text_props(c, gs[2]))

        @test [a.kind for a in XLSX.group_axes(c, gs[1])] == [:catAx, :valAx]
        @test [a.axid for a in XLSX.group_axes(c, gs[2])] == [1926562432, 1773317264]

        # The group is what ties a series to its axes.
        @test XLSX.series_group(c, 1).kind == :barChart
        @test XLSX.series_group(c, 3).kind == :lineChart
        @test [a.pos for a in XLSX.series_axes(c, 1)] == [:b, :l]
        @test [a.pos for a in XLSX.series_axes(c, 3)] == [:b, :r]
    end

    @testset "axes: identity and lookup" begin
        axes = XLSX.chart_axes(c)
        @test length(axes) == 4
        @test [a.kind for a in axes] == [:catAx, :valAx, :valAx, :catAx]
        @test [a.axid for a in axes] == [612078287, 460195247, 1773317264, 1926562432]
        @test [a.pos for a in axes] == [:b, :l, :r, :b]

        # A combo chart has two value axes, which is why this returns a vector.
        @test length(XLSX.chart_axes(c, :value)) == 2
        @test length(XLSX.chart_axes(c, :category)) == 2
        @test isempty(XLSX.chart_axes(c, :date))
        @test_throws XLSX.XLSXError XLSX.chart_axes(c, :nonsense)

        @test XLSX.chart_axis(c, 1773317264).pos == :r
        @test_throws XLSX.XLSXError XLSX.chart_axis(c, 1)

        # crossAx forms two closed pairs.
        @test XLSX.axis_partner(c, axes[1]).axid == 460195247
        @test XLSX.axis_partner(c, axes[2]).axid == 612078287
        @test XLSX.axis_partner(c, axes[3]).axid == 1926562432
        @test XLSX.axis_partner(c, axes[4]).axid == 1773317264
    end

    @testset "axes: deleted" begin
        del = XLSX.chart_axis(c, 1926562432)
        @test del.deleted
        @test all(!a.deleted for a in XLSX.chart_axes(c) if a.axid != 1926562432)

        # Excel strips formatting from a deleted axis but keeps its structure.
        @test isnothing(XLSX.axis_shape_props(c, del))
        @test isnothing(XLSX.axis_text_props(c, del))
        @test isnothing(XLSX.axis_title_text(c, del))
        @test isnothing(XLSX.axis_gridlines(c, del))
        @test XLSX.axis_label_offset(c, del) == 100      # scalars survive
        @test !isnothing(XLSX.axis_partner(c, del))
    end

    @testset "axes: appearance" begin
        cat, pri, sec = XLSX.chart_axis(c, 612078287),
                        XLSX.chart_axis(c, 460195247),
                        XLSX.chart_axis(c, 1773317264)

        @test XLSX.text_content(XLSX.axis_title_text(c, cat)) == "Horizontal"
        @test XLSX.text_content(XLSX.axis_title_text(c, pri)) == "Primary"
        @test XLSX.text_content(XLSX.axis_title_text(c, sec)) == "Secondary"

        # No fixture has a title bound to a cell.
        @test isnothing(XLSX.axis_title_ref(c, cat))

        # noFill plus a real line, versus noFill on both: same has_* answers
        # from different XML, and both explicit rather than absent.
        spc = XLSX.axis_shape_props(c, cat)
        @test !XLSX.has_fill(spc) && XLSX.has_line(spc)
        @test spc.fill.kind == :none                     # explicit <a:noFill/>

        spp = XLSX.axis_shape_props(c, pri)
        @test !XLSX.has_fill(spp) && !XLSX.has_line(spp)

        @test !isnothing(XLSX.axis_gridlines(c, pri))
        @test isnothing(XLSX.axis_gridlines(c, cat))
        @test isnothing(XLSX.axis_gridlines(c, pri; minor=true))
    end

    @testset "axes: scalars" begin
        cat, pri, sec = XLSX.chart_axis(c, 612078287),
                        XLSX.chart_axis(c, 460195247),
                        XLSX.chart_axis(c, 1773317264)

        @test XLSX.axis_number_format_code(c, cat) == "General"
        @test XLSX.axis_number_format_linked(c, cat) === true

        @test XLSX.axis_major_tick_mark(c, cat) == :none
        @test XLSX.axis_major_tick_mark(c, sec) == :out
        @test XLSX.axis_minor_tick_mark(c, cat) == :none
        @test XLSX.axis_tick_label_pos(c, cat) == :nextTo

        @test XLSX.axis_orientation(c, cat) == :minMax
        @test isnothing(XLSX.axis_min(c, pri))           # automatic scaling
        @test isnothing(XLSX.axis_max(c, pri))
        @test isnothing(XLSX.axis_log_base(c, pri))

        @test XLSX.axis_crosses(c, pri) == :autoZero
        @test XLSX.axis_crosses(c, sec) == :max
        @test isnothing(XLSX.axis_crosses_at(c, pri))    # mutually exclusive

        @test isnothing(XLSX.axis_major_unit(c, pri))
        @test isnothing(XLSX.axis_minor_unit(c, pri))

        @test XLSX.axis_label_align(c, cat) == :ctr
        @test XLSX.axis_label_offset(c, cat) == 100
        @test XLSX.axis_multi_level_labels(c, cat) === true   # noMultiLvlLbl="0"
        @test XLSX.axis_cross_between(c, pri) == :between

        # Kind-specific accessors throw rather than returning nothing, so that
        # nothing keeps meaning "not written".
        @test_throws XLSX.XLSXError XLSX.axis_cross_between(c, cat)
        @test_throws XLSX.XLSXError XLSX.axis_label_offset(c, pri)
        @test_throws XLSX.XLSXError XLSX.axis_label_align(c, pri)
        @test_throws XLSX.XLSXError XLSX.axis_major_unit(c, cat)
    end

    @testset "chart space level" begin
        # Title: txPr present, no c:tx at all — Excel generates the text.
        @test !isnothing(XLSX.chart_title_node(c))
        @test isnothing(XLSX.chart_title_text(c))
        @test isnothing(XLSX.chart_title_ref(c))
        @test XLSX.auto_title_deleted(c) === false
        ttp = XLSX.chart_title_text_props(c)
        @test XLSX.default_run_props(ttp).size == 14.0
        @test XLSX.default_run_props(ttp).fill.fgcolor.rgb == "595959"  # tx1 +lumMod/lumOff

        @test XLSX.legend_pos(c) == :b
        @test XLSX.legend_overlay(c) === false
        @test XLSX.default_run_props(XLSX.legend_text_props(c)).size == 9.0

        pa = XLSX.plotarea_shape_props(c)
        @test !XLSX.has_fill(pa) && !XLSX.has_line(pa)

        # The only element in the file with both a real fill and a real line.
        cs = XLSX.chartspace_shape_props(c)
        @test XLSX.has_fill(cs) && XLSX.has_line(cs)
        @test cs.fill.fgcolor.rgb == "FFFFFF"            # bg1 -> lt1
        @test cs.line.width == 0.75                      # w="9525" EMU -> points

        # The top of the cascade: present, but says nothing.
        cst = XLSX.chartspace_text_props(c)
        @test !isnothing(cst)
        @test isnothing(XLSX.default_run_props(cst).size)
    end

    @testset "data points" begin
        @test length(XLSX.series_data_points(c, 1)) == 1
        @test isempty(XLSX.series_data_points(c, 2))
        @test length(XLSX.series_data_points(c, 3)) == 1

        d1 = XLSX.series_data_point(c, 1, 2)             # 1-based position
        @test !isnothing(d1)
        @test d1.idx == 1                                # c:idx is 0-based
        @test d1.invert_if_negative === false            # written by Excel
        @test XLSX.data_point_shape_props(c, d1).fill.fgcolor.rgb == "FF0000"
        @test isnothing(XLSX.data_point_marker(c, d1))

        @test isnothing(XLSX.series_data_point(c, 1, 1))
        @test isnothing(XLSX.series_data_point(c, 1, 3))
        @test_throws XLSX.XLSXError XLSX.series_data_point(c, 1, 0)

        # A point that overrides only its marker: the point-level spPr is the
        # line segment, which has no fill; the colour lives on the marker.
        d3 = XLSX.series_data_point(c, 3, 3)
        @test d3.idx == 2
        @test isnothing(d3.invert_if_negative)           # meaningless on a line
        sp = XLSX.data_point_shape_props(c, d3)
        @test isnothing(sp.fill)                         # absent, not noFill
        @test sp.line.width == 2.25
        mk = XLSX.data_point_marker(c, d3)
        @test mk.symbol == :diamond && mk.size == 9
        @test mk.shape.fill.fgcolor.rgb == "FF0000"

        # idx round-trips through the 1-based lookup to the same node.
        for i in 1:length(c.series), d in XLSX.series_data_points(c, i)
            @test XLSX.series_data_point(c, i, d.idx + 1).raw === d.raw
        end
    end

    @testset "individual data labels" begin
        dls = XLSX.series_data_labels(c, 1)
        @test length(dls) == 3
        @test [d.idx for d in dls] == [0, 1, 2]
        @test isempty(XLSX.series_data_labels(c, 2))

        # Retyped label: literal text, and formatting in both c:rich and c:txPr.
        r = XLSX.series_data_label(c, 1, 1)
        @test XLSX.text_content(XLSX.data_label_text(c, r)) == "Best"
        @test !isnothing(XLSX.data_label_text_props(c, r))
        @test XLSX.default_run_props(XLSX.data_label_text(c, r)).fill.fgcolor.rgb == "00B0F0"
        @test isnothing(XLSX.data_label_offset(c, r))
        @test isnothing(XLSX.data_label_position(c, r))

        # Dragged label: a manual layout offset, no dLblPos.
        m = XLSX.series_data_label(c, 1, 2)
        off = XLSX.data_label_offset(c, m)
        @test off.x ≈ -0.038888888888888994
        @test off.y ≈ -0.06481481481481485
        @test isnothing(XLSX.data_label_position(c, m))
        @test isnothing(XLSX.data_label_text(c, m))

        # Deleted label: c:delete and nothing else.
        x = XLSX.series_data_label(c, 1, 3)
        @test x.delete === true
        @test isnothing(XLSX.data_label_text(c, x))
        @test isnothing(XLSX.data_label_text_props(c, x))

        # Absent delete means shown, and is distinct from an explicit false.
        @test r.delete === nothing
        @test m.delete === nothing
        @test length(filter(d -> d.delete !== true, dls)) == 2
    end

    @testset "trendlines" begin
        @test isempty(XLSX.series_trendlines(c, 2))
        ts = XLSX.series_trendlines(c, 1)
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

        sp = XLSX.trendline_shape_props(c, t)
        @test sp.line.width == 1.5                       # w="19050"
        @test sp.line.dash == "sysDot"                   # element, not attribute
        @test isnothing(sp.fill)

        # The label carries formatting but no typed text.
        @test isnothing(XLSX.trendline_label_text(c, t))
        @test !isnothing(XLSX.trendline_label_text_props(c, t))
        @test !isnothing(XLSX.trendline_label_shape_props(c, t))
    end

    @testset "error bars" begin
        @test isempty(XLSX.series_error_bars(c, 1))
        es = XLSX.series_error_bars(c, 2)
        @test length(es) == 1

        e = es[1]
        @test isnothing(e.direction)                     # omitted on a bar chart
        @test e.bar_type == :both
        @test e.value_type == :fixedVal
        @test e.value == 0.5
        @test e.no_end_cap === false

        sp = XLSX.error_bars_shape_props(c, e)
        @test XLSX.has_line(sp)
        @test sp.fill.kind == :none                      # explicit noFill

        # Only cust bars carry plus/minus references.
        @test XLSX.error_bars_custom_refs(c, e) == (plus = nothing, minus = nothing)
    end

    @testset "other fixtures" begin
        fb = XLSX.readxlsx(joinpath(data_directory, "chart_basic.xlsx"))
        cb = XLSX.getCharts(fb)[1]

        # A literal title, reached two ways: the accessor and parse_chart_title.
        @test XLSX.text_content(XLSX.chart_title_text(cb)) == "Revenue by Region"
        @test cb.title == "Revenue by Region"
        @test length(XLSX.chart_groups(cb)) == 1
        @test isempty(XLSX.series_data_points(cb, 1))

        # Six theme variants, all reached through the accessor.
        ft = XLSX.readxlsx(joinpath(data_directory, "chart_theme_colors.xlsx"))
        ct = XLSX.getCharts(ft)[1]
        @test length(ct.series) == 6
        @test XLSX.series_shape_props(ct, 1).fill.fgcolor.rgb == "156082"
        @test XLSX.series_shape_props(ct, 5).fill.fgcolor.rgb == "104862"

        # A cx: chart is a ChartEx, and the c: accessors are not defined for it.
        fm = XLSX.readxlsx(joinpath(data_directory, "chart_mixed.xlsx"))
        cms = XLSX.getCharts(fm)
        @test any(x -> x isa XLSX.ChartEx, cms)
        cx = first(x for x in cms if x isa XLSX.ChartEx)
        @test_throws MethodError XLSX.series_shape_props(cx, 1)
    end
end