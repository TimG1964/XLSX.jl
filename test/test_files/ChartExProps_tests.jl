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
end