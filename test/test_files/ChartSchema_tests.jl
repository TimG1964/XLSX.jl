# Schema tables and the node-insertion primitives they drive.
#
# Fixture: chart_appearance.xlsx — see the inventory at the top of
# ChartProps_tests.jl before assuming what it contains.

const NS_A                   = XLSX.NS_A
const NS_C                   = XLSX.NS_C
const XLSXError              = XLSX.XLSXError
const CHILD_ORDER            = XLSX.CHILD_ORDER
const _slot                  = XLSX._slot
const insert_child           = XLSX.insert_child
const replace_child          = XLSX.replace_child
const rebuild_path           = XLSX.rebuild_path
const ns_prefixes            = XLSX.ns_prefixes
const prefixed_tag           = XLSX.prefixed_tag
const localname              = XLSX.localname
const first_element_with_tag = XLSX.first_element_with_tag
const _series                = XLSX._series
const getChartAxis           = XLSX.getChartAxis
const remove_child           = XLSX.remove_child
const remove_choice          = XLSX.remove_choice

@testset "ChartSchema" begin

    xf = XLSX.openxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
    c = first(XLSX.getCharts(xf[1]))
    root = XLSX.chart_root(c)
    pfx = ns_prefixes(root)
    mk(ns, tag; kw...) = XML.Element(prefixed_tag(pfx[ns], tag); kw...)
    SP = (NS_A, "spPr")

    @testset "schema tables" begin
        for (key, order) in CHILD_ORDER
            @test !isempty(order)
            @test key[1] in (NS_C, NS_A)
        end
        pa = CHILD_ORDER[(NS_C, "plotArea")]
        @test _slot(pa, "catAx") == _slot(pa, "valAx") == _slot(pa, "dateAx")
        @test _slot(pa, "layout") < _slot(pa, "barChart") < _slot(pa, "catAx") < _slot(pa, "spPr")
        @test isnothing(_slot(pa, "nonsense"))
    end

    @testset "insert_child" begin
        sp3 = first_element_with_tag(_series(c, 3).raw, "spPr")
        sp3b = insert_child(sp3, SP, mk(NS_A, "noFill"))
        @test localname.(collect(XML.eachelement(sp3b))) == ["noFill", "ln", "effectLst"]
        @test localname.(collect(XML.eachelement(sp3))) == ["ln", "effectLst"]

        @test occursin("noFill", XML.write(insert_child(mk(NS_C, "spPr"), SP, mk(NS_A, "noFill"))))

        # A PARSED childless element has `children === nothing`, unlike a constructed
        # one. This is the case that makes insert_child return a node rather than mutate.
        parsed = XML.parse("""<c:spPr xmlns:c="$NS_C"/>""", XML.Node)[1]
        @test isnothing(parsed.children)
        @test occursin("noFill", XML.write(insert_child(parsed, SP, mk(NS_A, "noFill"))))

        bare = XML.parse("""<c:ser xmlns:c="$NS_C"><c:idx val="0"/><c:order val="0"/><c:val/></c:ser>""", XML.Node)[1]
        got = insert_child(bare, (NS_C, "barSer"), mk(NS_C, "spPr"))
        @test localname.(collect(XML.eachelement(got))) == ["idx", "order", "spPr", "val"]

        twice = insert_child(got, (NS_C, "barSer"), mk(NS_C, "spPr"))
        @test count(k -> localname(k) == "spPr", XML.children(twice)) == 1

        @test_throws XLSXError insert_child(bare, (NS_C, "nonesuch"), mk(NS_C, "spPr"))
        @test_throws XLSXError insert_child(bare, (NS_C, "barSer"), mk(NS_C, "axId"))
    end

    @testset "xsd:choice" begin
        ax = getChartAxis(c, 1773317264).raw          # has crosses="max"
        @test_throws XLSXError insert_child(ax, (NS_C, "valAx"), mk(NS_C, "crossesAt"; val="0"))

        sp = insert_child(first_element_with_tag(_series(c, 3).raw, "spPr"), SP, mk(NS_A, "noFill"))
        @test_throws XLSXError insert_child(sp, SP, mk(NS_A, "solidFill"))
        # A member does not exclude itself: replacing a fill with the same kind is fine.
        @test !isnothing(insert_child(sp, SP, mk(NS_A, "noFill")))
    end

    @testset "rebuild_path" begin
        ser3 = _series(c, 3).raw
        newroot = rebuild_path(root,
            [(NS_C, "chart") => "chart",
                (NS_C, "plotArea") => "plotArea",
                (NS_C, "lineChart") => "lineChart",
                (NS_C, "lineSer") => ("ser", n -> n === ser3),
                SP => "spPr"],
            sp -> insert_child(sp, SP, mk(NS_A, "noFill"));
            prefixes=pfx)

        @test newroot !== root
        @test first_element_with_tag(newroot, "chart") !== first_element_with_tag(root, "chart")
        lc = first_element_with_tag(first_element_with_tag(
            first_element_with_tag(newroot, "chart"), "plotArea"), "lineChart")
        @test occursin("noFill", XML.write(first_element_with_tag(
            first_element_with_tag(lc, "ser"), "spPr")))

        # A predicate that matches nothing throws rather than creating.
        @test_throws XLSXError rebuild_path(root,
            [(NS_C, "chart") => "chart",
                (NS_C, "plotArea") => "plotArea",
                (NS_C, "lineChart") => ("lineChart", n -> false)],
            identity; prefixes=pfx)
    end

    @testset "remove_child" begin
        bare = XML.parse("""<c:ser xmlns:c="$NS_C"><c:idx val="0"/><c:order val="0"/><c:val/></c:ser>""", XML.Node)[1]

        got = remove_child(bare, "order")
        @test localname.(collect(XML.eachelement(got))) == ["idx", "val"]
        @test localname.(collect(XML.eachelement(bare))) == ["idx", "order", "val"]  # original untouched

        # absent child is a no-op, and returns the same object
        @test remove_child(bare, "spPr") === bare

        # childless parent is a no-op
        parsed = XML.parse("""<c:spPr xmlns:c="$NS_C"/>""", XML.Node)[1]
        @test remove_child(parsed, "ln") === parsed

        # only the first match goes
        two = XML.parse("""<c:ser xmlns:c="$NS_C"><c:dPt/><c:dPt/></c:ser>""", XML.Node)[1]
        @test length(XML.children(remove_child(two, "dPt"))) == 1
    end

    @testset "remove_choice" begin
        SP = (NS_A, "spPr")

        # removes whichever member is present, not the one named
        sp = insert_child(XML.parse("""<c:spPr xmlns:c="$NS_C"/>""", XML.Node)[1],
            SP, mk(NS_A, "noFill"))
        @test localname.(collect(XML.eachelement(sp))) == ["noFill"]
        @test isempty(XML.children(remove_choice(sp, SP, "solidFill")))

        # leaves other slots alone: a line survives fill removal
        sp2 = insert_child(insert_child(XML.parse("""<c:spPr xmlns:c="$NS_C"/>""", XML.Node)[1],
            SP, mk(NS_A, "noFill")),
            SP, mk(NS_A, "ln"))
        @test localname.(collect(XML.eachelement(remove_choice(sp2, SP, "solidFill")))) == ["ln"]

        # nothing present is a no-op
        lnonly = insert_child(XML.parse("""<c:spPr xmlns:c="$NS_C"/>""", XML.Node)[1],
            SP, mk(NS_A, "ln"))
        @test localname.(collect(XML.eachelement(remove_choice(lnonly, SP, "noFill")))) == ["ln"]

        # a tag that is not in any choice group throws
        @test_throws XLSXError remove_choice(sp, SP, "ln")
        @test_throws XLSXError remove_choice(sp, SP, "nonsense")
    end

end # ChartSchema