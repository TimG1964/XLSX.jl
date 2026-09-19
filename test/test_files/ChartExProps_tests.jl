@testset "ChartExProps" begin
    @testset "series, layout and data blocks" begin
        f  = XLSX.readxlsx(joinpath(data_directory, "chart_ex.xlsx"))
        cx = only(filter(x -> x isa XLSX.ChartEx, XLSX.getCharts(f)))

        @test XLSX.getChartSeriesCount(cx) == 1
        @test XLSX.getSeriesLayout(cx, 1) == :waterfall
        @test_throws XLSX.XLSXError XLSX.getSeriesLayout(cx, 2)

        blocks = XLSX.getChartDataBlocks(cx)
        @test length(blocks) == 1
        b = only(blocks)
        @test b.id == 0
        @test [d.kind for d in b.dimensions] == [:str, :num]
        @test [d.type for d in b.dimensions] == [:cat, :val]
        @test all(d -> startswith(d.formula, "_xlchart."), b.dimensions)
        @test all(d -> d.range isa XLSX.SheetCellRange, b.dimensions)

        @test XLSX.getSeriesData(cx, 1) == b
        @test isnothing(XLSX.getSeriesName(cx, 1))
        @test isnothing(XLSX.getSeriesNameRange(cx, 1))
    end

    @testset "layouts fixture" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chartex_layouts.xlsx"))
        cxs = filter(x -> x isa XLSX.ChartEx, XLSX.getCharts(f))
        chart(sheet) = only(filter(x -> x.sheet == sheet, cxs))

        @testset "waterfall" begin
            c = chart("waterfall")
            @test XLSX.getSeriesLayout(c, 1) == :waterfall
            @test XLSX.getSeriesSubtotals(c, 1) == [2, 5]
            @test XLSX.getLabelPosition(c, 1) == :outEnd
            @test XLSX.getLabelFlag(c, 1, :value) === true
            @test XLSX.getLabelFlag(c, 1, :seriesName) === false
            @test_throws XLSX.XLSXError XLSX.getLabelFlag(c, 1, :nonsense)
            @test isnothing(XLSX.charttitle(c))
            @test isnothing(XLSX.getSeriesName(c, 1))
        end

        @testset "funnel" begin
            c = chart("funnel")
            @test XLSX.getSeriesLayout(c, 1) == :funnel
            @test only(only(XLSX.getChartDataBlocks(c)).dimensions).formula == "_xlchart.v2.3"
            @test isnothing(XLSX.getSeriesSubtotals(c, 1))
            @test isnothing(XLSX.getLabelPosition(c, 1))
            @test isnothing(XLSX.getLabelFlag(c, 1, :value))
        end

        @testset "treemap and sunburst" begin
            t = chart("treemap")
            @test [d.type for d in only(XLSX.getChartDataBlocks(t)).dimensions] == [:cat, :size]
            @test XLSX.getSeriesName(t, 1) == "Value"
            @test !isnothing(XLSX.getSeriesNameRange(t, 1))
            @test XLSX.getSeriesParentLabelLayout(t, 1) == :overlapping
            @test XLSX.getLabelPosition(t, 1) == :inEnd
            @test XLSX.getLabelFlag(t, 1, :categoryName) === true

            s = chart("sunburst")
            @test XLSX.getSeriesLayout(s, 1) == :sunburst
            @test isnothing(XLSX.getSeriesParentLabelLayout(s, 1))   # no cx:layoutPr
            @test XLSX.getLabelPosition(s, 1) == :ctr
        end

        @testset "histogram" begin
            c = chart("histogram")
            @test XLSX.chartType(c) == :histogram                    # binning read inside layoutPr
            @test XLSX.getSeriesLayout(c, 1) == :clusteredColumn
            @test XLSX.getSeriesBinning(c, 1) == XLSX.ChartExBinning(:r, 0.0, 100.0, 10.0, nothing)
            @test XLSX.charttitle(c) == "Chart Title"                 # typed: txData without f
            @test isnothing(XLSX.getChartTitleRange(c))
            @test XLSX.getSeriesName(c, 1) == "Values"
        end

        @testset "pareto" begin
            c = chart("pareto")
            @test XLSX.getChartSeriesCount(c) == 4
            @test [XLSX.getSeriesLayout(c, i) for i in 1:4] ==
                [:clusteredColumn, :paretoLine, :clusteredColumn, :paretoLine]
            @test [XLSX.getSeriesOwner(c, i) for i in 1:4] == [nothing, 1, nothing, 3]
            @test [XLSX.getSeriesHidden(c, i) for i in 1:4] == [nothing, nothing, true, nothing]
            @test [XLSX.getSeriesAxisIds(c, i) for i in 1:4] == [[1], [2], [1], [2]]
            @test [XLSX.getSeriesAggregation(c, i) for i in 1:4] == [true, false, true, false]
            @test isnothing(XLSX.getSeriesData(c, 2))                # a paretoLine has no dataId
            blocks = XLSX.getChartDataBlocks(c)
            @test length(blocks) == 2
            @test blocks[1].dimensions[1].formula == blocks[2].dimensions[1].formula  # shared name
            @test XLSX.getSeriesData(c, 3) == blocks[2]
            @test XLSX.chartType(c) == :pareto
        end

        @testset "box and whisker" begin
            c = chart("boxwhisker")
            @test XLSX.getChartSeriesCount(c) == 3
            @test [XLSX.getSeriesName(c, i) for i in 1:3] == ["Group A", "Group B", "Group C"]
            @test [XLSX.getSeriesQuartileMethod(c, i) for i in 1:3] == [:exclusive, :inclusive, :exclusive]
            @test [XLSX.getSeriesLayoutFlag(c, i, :meanMarker) for i in 1:3] == [true, false, true]
            @test [XLSX.getSeriesLayoutFlag(c, i, :outliers) for i in 1:3] == [true, false, true]
            @test all(i -> XLSX.getSeriesLayoutFlag(c, i, :meanLine) === false, 1:3)
            @test [XLSX.getSeriesData(c, i).id for i in 1:3] == [0, 1, 2]
            @test XLSX.chartType(c) == :boxWhisker
        end

        @testset "bound title and series name" begin
            c = chart("bound")
            @test XLSX.charttitle(c) == "Title here"
            @test !isnothing(XLSX.getChartTitleRange(c))
            @test XLSX.getSeriesName(c, 1) == "Series name"
            @test !isnothing(XLSX.getSeriesNameRange(c, 1))
            @test XLSX.getSeriesSubtotals(c, 1) == Int[]              # present, empty
        end
    end

    @testset "formatting cascade" begin
        f = XLSX.readxlsx(joinpath(data_directory, "chartex_formatted.xlsx"))
        c = only(filter(x -> x isa XLSX.ChartEx, XLSX.getCharts(f)))

        # Series fill: theme colour with a transform.
        e = XLSX.getSeriesFill(c, 1)
        @test e.site.level == :series
        @test e.value.fgcolor.val == "accent2"

        # Point 5 (idx 4) has its own fill: explicit colour with transforms.
        e5 = XLSX.getSeriesFill(c, 1; point = 5)
        @test e5.site.level == :point
        @test [s.level for s in e5.chain] == [:point, :series]
        @test e5.value.fgcolor.kind == :srgb
        @test e5.value.fgcolor.val == "4EA72E"                                  # as written
        @test e5.value.fgcolor.transforms == [:lumMod => 60000, :lumOff => 40000]
        @test e5.value.fgcolor.rgb == "8ED973"                                  # transforms applied

        # Point 1 has no dataPt: no point rung, resolves at the series.
        e1 = XLSX.getSeriesFill(c, 1; point = 1)
        @test e1.site.level == :series
        @test [s.level for s in e1.chain] == [:series]
        @test_throws XLSX.XLSXError XLSX.getSeriesFill(c, 1; point = 0)

        # Label text: set on dataLabels/txPr defRPr.
        s = XLSX.getLabelTextProp(c, 1, :size)
        @test s.value ≈ 11.0
        @test s.site.level == :series
        @test [x.level for x in s.chain] == [:series, :chartspace]
        @test XLSX.getLabelTextProp(c, 1, :bold).value === true
        @test XLSX.getLabelTextProp(c, 1, :fill).value.fgcolor.val == "7030A0"
        @test_throws XLSX.XLSXError XLSX.getLabelTextProp(c, 1, :nonsense)

        # Title: formatting on the run, not defRPr.
        tp = XLSX.default_run_props(XLSX.getChartTitleTextProps(c))
        @test tp.size ≈ 18.0
        @test tp.latin == "Comic Sans MS"

        # Legend: no run; defRPr sets only the colour, size is on endParaRPr.
        lg = XLSX.default_run_props(XLSX.getLegendTextProps(c))
        @test lg.fill.fgcolor.val == "accent3"
        @test isnothing(lg.size)

        # Nothing written anywhere: no value, no site, but a full chain.
        g = XLSX.readxlsx(joinpath(data_directory, "chartex_layouts.xlsx"))
        w = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "waterfall", XLSX.getCharts(g)))
        u = XLSX.getSeriesFill(w, 1)
        @test isnothing(u.value) && isnothing(u.site)
        @test length(u.chain) == 1
        @test isnothing(only(u.chain).props)
        @test isnothing(XLSX.getLabelTextProp(w, 1, :size).value)

    end

    @testset "setSeriesFill" begin
        tmp = joinpath(mktempdir(), "cx_fill.xlsx")
        cp(joinpath(data_directory, "chartex_formatted.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode = "rw")
        c = only(filter(x -> x isa XLSX.ChartEx, XLSX.getCharts(xf)))

        # series fill, replacing a theme colour
        @test XLSX.setSeriesFill(c, 1, "red") === c
        @test XLSX.getSeriesFill(c, 1).value.fgcolor.rgb == "FF0000"   # same handle, still valid

        # an existing dataPt (point 5, idx 4) is edited in place
        XLSX.setSeriesFill(c, 1, XLSX.SchemeColor(:accent1); point = 5)
        @test XLSX.getSeriesFill(c, 1; point = 5).value.fgcolor.val == "accent1"

        # new dataPts: one before idx 4, one after, kept in idx order
        XLSX.setSeriesFill(c, 1, "blue";  point = 2)
        XLSX.setSeriesFill(c, 1, "green"; point = 7)
        ser = XLSX._cx_series_node(c, 1)
        @test [XLSX._cx_idx(n) for n in XLSX.elements_with_tag(ser, "dataPt")] == [1, 4, 6]
        @test XLSX.getSeriesFill(c, 1; point = 2).site.level == :point

        # child order of the series follows the schema
        @test [String(XLSX.localname(n)) for n in XLSX.XML.eachelement(ser)] ==
              ["spPr", "dataPt", "dataPt", "dataPt", "dataLabels", "dataId", "layoutPr"]

        # :none is an explicit setting; :inherit on an unformatted point creates nothing
        XLSX.setSeriesFill(c, 1, :none; point = 2)
        @test XLSX.getSeriesFill(c, 1; point = 2).value.kind === :none
        XLSX.setSeriesFill(c, 1, :inherit; point = 3)
        @test length(XLSX.elements_with_tag(XLSX._cx_series_node(c, 1), "dataPt")) == 3

        @test_throws XLSX.XLSXError XLSX.setSeriesFill(c, 1, "red"; point = 0)
        @test_throws XLSX.XLSXError XLSX.setSeriesFill(c, 2, "red")   # only one series

        @test all(n -> XLSX.XML.tag(XLSX.first_element_with_tag(n, "spPr")) == "cx:spPr",
            XLSX.elements_with_tag(XLSX._cx_series_node(c, 1), "dataPt"))

        # the changes reach disk
        out = joinpath(mktempdir(), "cx_fill_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite = true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.ChartEx, XLSX.getCharts(f)))
        @test XLSX.getSeriesFill(d, 1).value.fgcolor.rgb == "FF0000"
        @test XLSX.getSeriesFill(d, 1; point = 7).value.fgcolor.rgb == "008000"
        @test XLSX.getSeriesFill(d, 1; point = 2).value.kind === :none
    end

    @testset "setSeriesFill :inherit writes nothing where nothing was set" begin
        tmp = joinpath(mktempdir(), "cx_inherit.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode = "rw")
        w = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "waterfall", XLSX.getCharts(xf)))
        XLSX.setSeriesFill(w, 1, :inherit)
        @test isnothing(XLSX.first_element_with_tag(XLSX._cx_series_node(w, 1), "spPr"))
    end

    @testset "setSeriesLine" begin
        tmp = joinpath(mktempdir(), "cx_line.xlsx")
        cp(joinpath(data_directory, "chartex_formatted.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode = "rw")
        c = only(filter(x -> x isa XLSX.ChartEx, XLSX.getCharts(xf)))

        # several properties at once, on a series with no a:ln
        XLSX.setSeriesLine(c, 1; color = "FFC00000", width = 2.25, dash = :dash, cap = :round)
        l = XLSX.getSeriesLine(c, 1).value
        @test l.fill.fgcolor.rgb == "C00000"
        @test l.width ≈ 2.25
        @test l.dash == "dash"
        @test l.cap == "rnd"

        # an Excel UI alias resolves to the DrawingML spelling
        XLSX.setSeriesLineDash(c, 1, :roundDot)
        @test XLSX.getSeriesLine(c, 1).value.dash == "sysDot"

        # one property back to inherit, leaving the rest
        XLSX.setSeriesLine(c, 1; width = :inherit)
        l = XLSX.getSeriesLine(c, 1).value
        @test isnothing(l.width) && l.fill.fgcolor.rgb == "C00000"

        # a created point gets its own outline
        XLSX.setSeriesLine(c, 1; point = 2, color = "FF0070C0", width = 1.5)
        @test XLSX.getSeriesLine(c, 1; point = 2).site.level == :point
        dp = only(filter(n -> XLSX._cx_idx(n) == 1,
                         XLSX.elements_with_tag(XLSX._cx_series_node(c, 1), "dataPt")))
        @test XLSX.XML.tag(XLSX.first_element_with_tag(dp, "spPr")) == "cx:spPr"

        # :none is explicit, :inherit removes the a:ln
        XLSX.setSeriesLine(c, 1, :none)
        @test XLSX.getSeriesLine(c, 1).value.fill.kind === :none
        XLSX.setSeriesLine(c, 1, :inherit)
        @test isnothing(XLSX.first_element_with_tag(
                  XLSX.first_element_with_tag(XLSX._cx_series_node(c, 1), "spPr"), "ln"))

        # no keywords is a no-op; a bad symbol throws
        @test XLSX.setSeriesLine(c, 1) === c
        @test_throws XLSX.XLSXError XLSX.setSeriesLine(c, 1, :dotted)

        # reaches disk
        out = joinpath(mktempdir(), "cx_line_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite = true)
        d = only(filter(x -> x isa XLSX.ChartEx, XLSX.getCharts(XLSX.readxlsx(out))))
        @test XLSX.getSeriesLine(d, 1; point = 2).value.width ≈ 1.5
    end

    @testset "setLabelTextProp" begin
        tmp = joinpath(mktempdir(), "cx_label.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode = "rw")
        # the funnel has no cx:dataLabels at all
        c = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "funnel", XLSX.getCharts(xf)))

        @test isnothing(XLSX.getLabelTextProp(c, 1, :size).value)
        XLSX.setLabelTextProp(c, 1, :size, 11)
        @test XLSX.getLabelTextProp(c, 1, :size).value ≈ 11.0

        # a created cx:dataLabels must not turn labels on
        lbl = XLSX.first_element_with_tag(XLSX._cx_series_node(c, 1), "dataLabels")
        @test XLSX.XML.tag(lbl) == "cx:dataLabels"
        @test [String(XLSX.localname(n)) for n in XLSX.XML.eachelement(lbl)] == ["txPr", "visibility"]
        @test all(f -> XLSX.getLabelFlag(c, 1, f) === false,
                  (:seriesName, :categoryName, :value))
        @test XLSX.XML.tag(XLSX.first_element_with_tag(lbl, "txPr")) == "cx:txPr"

        # more fields, and a composite one
        XLSX.setLabelTextProp(c, 1, :bold, true)
        XLSX.setLabelTextProp(c, 1, :fill, "FF7030A0")
        @test XLSX.getLabelTextProp(c, 1, :bold).value === true
        @test XLSX.getLabelTextProp(c, 1, :fill).value.fgcolor.rgb == "7030A0"
        @test XLSX.getLabelTextProp(c, 1, :size).value ≈ 11.0     # earlier field kept

        XLSX.setLabelTextProp(c, 1, :size, :inherit)
        @test isnothing(XLSX.getLabelTextProp(c, 1, :size).value)
        @test_throws XLSX.XLSXError XLSX.setLabelTextProp(c, 1, :nonsense, 1)

        # an existing cx:dataLabels keeps the flags Excel wrote
        w = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "waterfall", XLSX.getCharts(xf)))
        XLSX.setLabelTextProp(w, 1, :size, 9)
        @test XLSX.getLabelFlag(w, 1, :value) === true
        @test XLSX.getLabelTextProp(w, 1, :size).value ≈ 9.0

        out = joinpath(mktempdir(), "cx_label_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite = true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "funnel", XLSX.getCharts(f)))
        @test XLSX.getLabelTextProp(d, 1, :bold).value === true
    end

    @testset "title, legend and axis text" begin
        tmp = joinpath(mktempdir(), "cx_text.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode = "rw")
        w = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "waterfall", XLSX.getCharts(xf)))

        # title: cx:title exists but has no txPr
        XLSX.setChartTitleTextProp(w, :size, 16)
        XLSX.setChartTitleTextProp(w, :fill, "FF7030A0")
        tp = XLSX.default_run_props(XLSX.getChartTitleTextProps(w))
        @test tp.size ≈ 16.0 && tp.fill.fgcolor.rgb == "7030A0"
        @test XLSX.XML.tag(XLSX.first_element_with_tag(
                  XLSX.first_element_with_tag(XLSX._cx_chart(w), "title"), "txPr")) == "cx:txPr"

        # legend
        XLSX.setLegendTextProp(w, :bold, true)
        @test XLSX.default_run_props(XLSX.getLegendTextProps(w)).bold === true

        # axes: by id, not position
        @test XLSX.getChartAxisIds(w) == [0, 1]
        XLSX.setAxisTitleTextProp(w, 1, :size, 12)
        @test XLSX.default_run_props(XLSX.getAxisTitleTextProps(w, 1)).size ≈ 12.0
        @test isnothing(XLSX.getAxisTitleTextProps(w, 0))          # untouched axis
        @test_throws XLSX.XLSXError XLSX.setAxisTitleTextProp(w, 7, :size, 12)

        # the legend is created where there is none (the funnel has no cx:legend)
        fn = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "funnel", XLSX.getCharts(xf)))
        XLSX.setLegendTextProp(fn, :size, 9)
        lg = XLSX.first_element_with_tag(XLSX._cx_chart(fn), "legend")
        @test XLSX.XML.tag(lg) == "cx:legend"
        @test XLSX.default_run_props(XLSX.getLegendTextProps(fn)).size ≈ 9.0

        out = joinpath(mktempdir(), "cx_text_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite = true)
        d = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "waterfall",
                        XLSX.getCharts(XLSX.readxlsx(out))))
        @test XLSX.default_run_props(XLSX.getChartTitleTextProps(d)).size ≈ 16.0
    end

    @testset "setChartTitleText and setSeriesName" begin
        tmp = joinpath(mktempdir(), "cx_text_set.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode = "rw")

        # histogram: a typed title, with the text also in the txPr runs
        h = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "histogram", XLSX.getCharts(xf)))
        @test XLSX.charttitle(h) == "Chart Title"
        XLSX.setChartTitleText(h, "Distribution of values")
        @test XLSX.charttitle(h) == "Distribution of values"
        @test XLSX.text_content(XLSX.getChartTitleTextProps(h)) == "Distribution of values"
        @test XLSX.default_run_props(XLSX.getChartTitleTextProps(h)).size ≈ 14.0  # formatting kept

        # bound: typing replaces the binding
        b = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "bound", XLSX.getCharts(xf)))
        @test !isnothing(XLSX.getChartTitleRange(b))
        XLSX.setChartTitleText(b, "Typed over")
        @test XLSX.charttitle(b) == "Typed over"
        @test isnothing(XLSX.getChartTitleRange(b))

        # waterfall: no title text at all, so this creates it
        w = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "waterfall", XLSX.getCharts(xf)))
        @test isnothing(XLSX.charttitle(w))
        XLSX.setChartTitleText(w, "Cash flow")
        @test XLSX.charttitle(w) == "Cash flow"

        # series names
        @test XLSX.getSeriesName(b, 1) == "Series name"
        XLSX.setSeriesName(b, 1, "Renamed")
        @test XLSX.getSeriesName(b, 1) == "Renamed"
        @test isnothing(XLSX.getSeriesNameRange(b, 1))
        @test isnothing(XLSX.getSeriesName(w, 1))
        XLSX.setSeriesName(w, 1, "Movement")
        @test XLSX.getSeriesName(w, 1) == "Movement"

        out = joinpath(mktempdir(), "cx_text_set_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite = true)
        f = XLSX.readxlsx(out)
        d = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "histogram", XLSX.getCharts(f)))
        @test XLSX.charttitle(d) == "Distribution of values"
    end

    @testset "setSeriesSubtotals" begin
        tmp = joinpath(mktempdir(), "cx_subtotals.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode = "rw")
        w = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "waterfall", XLSX.getCharts(xf)))

        @test XLSX.getSeriesSubtotals(w, 1) == [2, 5]     # what Excel wrote
        XLSX.setSeriesSubtotals(w, 1, [5, 3])
        @test XLSX.getSeriesSubtotals(w, 1) == [3, 5]     # sorted, replaced not appended
        XLSX.setSeriesSubtotals(w, 1, Int[])
        @test XLSX.getSeriesSubtotals(w, 1) == Int[]
        @test !isnothing(XLSX.first_element_with_tag(
                  XLSX._cx_layoutpr(w, 1), "subtotals"))  # present but empty
        XLSX.setSeriesSubtotals(w, 1, :inherit)
        @test isnothing(XLSX.getSeriesSubtotals(w, 1))
        @test_throws XLSX.XLSXError XLSX.setSeriesSubtotals(w, 1, [0, 2])

        # a series with no cx:layoutPr at all: the funnel
        fn = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "funnel", XLSX.getCharts(xf)))
        @test isnothing(XLSX._cx_layoutpr(fn, 1))
        XLSX.setSeriesSubtotals(fn, 1, [2])
        @test XLSX.getSeriesSubtotals(fn, 1) == [2]
        @test XLSX.XML.tag(XLSX._cx_layoutpr(fn, 1)) == "cx:layoutPr"

        out = joinpath(mktempdir(), "cx_subtotals_out.xlsx")
        XLSX.writexlsx(out, xf, overwrite = true)
        d = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "waterfall",
                        XLSX.getCharts(XLSX.readxlsx(out))))
        @test isnothing(XLSX.getSeriesSubtotals(d, 1))
    end

    @testset "layoutPr setters" begin
        tmp = joinpath(mktempdir(), "cx_layoutpr.xlsx")
        cp(joinpath(data_directory, "chartex_layouts.xlsx"), tmp)
        xf = XLSX.openxlsx(tmp; mode = "rw")
        chart(sheet) = only(filter(x -> x isa XLSX.ChartEx && x.sheet == sheet,
                                   XLSX.getCharts(xf)))

        @testset "quartile method and parent label layout" begin
            b = chart("boxwhisker")
            @test XLSX.getSeriesQuartileMethod(b, 1) == :exclusive
            XLSX.setSeriesQuartileMethod(b, 1, :inclusive)
            @test XLSX.getSeriesQuartileMethod(b, 1) == :inclusive
            XLSX.setSeriesQuartileMethod(b, 1, :inherit)
            @test isnothing(XLSX.getSeriesQuartileMethod(b, 1))
            @test_throws XLSX.XLSXError XLSX.setSeriesQuartileMethod(b, 1, :median)

            t = chart("treemap")
            @test XLSX.getSeriesParentLabelLayout(t, 1) == :overlapping
            XLSX.setSeriesParentLabelLayout(t, 1, :banner)
            @test XLSX.getSeriesParentLabelLayout(t, 1) == :banner
            @test_throws XLSX.XLSXError XLSX.setSeriesParentLabelLayout(t, 1, :sideways)

            # a series with no cx:layoutPr: the sunburst
            s = chart("sunburst")
            @test isnothing(XLSX._cx_layoutpr(s, 1))
            XLSX.setSeriesParentLabelLayout(s, 1, :none)
            @test XLSX.getSeriesParentLabelLayout(s, 1) == :none
        end

        @testset "layout flags" begin
            b = chart("boxwhisker")
            @test XLSX.getSeriesLayoutFlag(b, 2, :outliers) === false
            XLSX.setSeriesLayoutFlag(b, 2, :outliers, true)
            @test XLSX.getSeriesLayoutFlag(b, 2, :outliers) === true
            # the other attributes on the same cx:visibility are untouched
            @test XLSX.getSeriesLayoutFlag(b, 2, :meanMarker) === false
            XLSX.setSeriesLayoutFlag(b, 2, :outliers, :inherit)
            @test isnothing(XLSX.getSeriesLayoutFlag(b, 2, :outliers))
            @test XLSX.getSeriesLayoutFlag(b, 2, :meanMarker) === false

            # a waterfall flag on a series whose cx:visibility does not exist
            w = chart("waterfall")
            @test isnothing(XLSX.getSeriesLayoutFlag(w, 1, :connectorLines))
            XLSX.setSeriesLayoutFlag(w, 1, :connectorLines, false)
            @test XLSX.getSeriesLayoutFlag(w, 1, :connectorLines) === false
            # removing the only attribute removes the element
            XLSX.setSeriesLayoutFlag(w, 1, :connectorLines, :inherit)
            @test isnothing(XLSX.first_element_with_tag(XLSX._cx_layoutpr(w, 1), "visibility"))

            @test_throws XLSX.XLSXError XLSX.setSeriesLayoutFlag(w, 1, :nonsense, true)
            @test_throws XLSX.XLSXError XLSX.setSeriesLayoutFlag(w, 1, :connectorLines, 1)
        end

        @testset "binning and aggregation" begin
            h = chart("histogram")
            @test XLSX.getSeriesBinning(h, 1) ==
                  XLSX.ChartExBinning(:r, 0.0, 100.0, 10.0, nothing)

            XLSX.setSeriesBinning(h, 1; binSize = 20, overflow = 80, intervalClosed = :l)
            @test XLSX.getSeriesBinning(h, 1) ==
                  XLSX.ChartExBinning(:l, 0.0, 80.0, 20.0, nothing)

            # binCount replaces binSize
            XLSX.setSeriesBinning(h, 1; binCount = 7)
            b = XLSX.getSeriesBinning(h, 1)
            @test isnothing(b.binSize) && b.binCount == 7

            XLSX.setSeriesBinning(h, 1; underflow = :auto, overflow = :inherit)
            b = XLSX.getSeriesBinning(h, 1)
            @test b.underflow === :auto && isnothing(b.overflow)

            @test XLSX.setSeriesBinning(h, 1) === h        # no keywords is a no-op
            @test_throws XLSX.XLSXError XLSX.setSeriesBinning(h, 1; binSize = 5, binCount = 5)
            @test_throws XLSX.XLSXError XLSX.setSeriesBinning(h, 1; intervalClosed = :x)

            # aggregation and binning exclude each other
            p = chart("pareto")
            @test XLSX.getSeriesAggregation(p, 1) === true
            XLSX.setSeriesBinning(p, 1; binSize = 5)
            @test XLSX.getSeriesAggregation(p, 1) === false
            XLSX.setSeriesAggregation(p, 1, true)
            @test isnothing(XLSX.getSeriesBinning(p, 1))
            XLSX.setSeriesAggregation(p, 1, false)
            @test XLSX.getSeriesAggregation(p, 1) === false
        end

        @testset "reaches disk" begin
            out = joinpath(mktempdir(), "cx_layoutpr_out.xlsx")
            XLSX.writexlsx(out, xf, overwrite = true)
            f = XLSX.readxlsx(out)
            h = only(filter(x -> x isa XLSX.ChartEx && x.sheet == "histogram",
                            XLSX.getCharts(f)))
            @test XLSX.getSeriesBinning(h, 1).binCount == 7
        end
    end
end