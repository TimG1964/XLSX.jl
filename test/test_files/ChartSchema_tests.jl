# Schema tables and the node-insertion primitives they drive.
#
# Fixture: chart_appearance.xlsx — see the inventory at the top of
# ChartProps_tests.jl before assuming what it contains.


@testset "ChartSchema" begin

    xf = XLSX.openxlsx(joinpath(data_directory, "chart_appearance.xlsx"))
    c = first(XLSX.getCharts(xf[1]))
    root = XLSX.chart_root(c)
    pfx = XLSX.ns_prefixes(root)
    mk(ns, tag; kw...) = XML.Element(XLSX.prefixed_tag(pfx[ns], tag); kw...)
    SP = (XLSX.NS_A, "spPr")

    @testset "schema tables" begin
        for (key, order) in XLSX.CHILD_ORDER
            @test !isempty(order)
        @test key[1] in (XLSX.NS_C, XLSX.NS_A, XLSX.NS_CX)
        end
        pa = XLSX.CHILD_ORDER[(XLSX.NS_C, "plotArea")]
        @test XLSX._slot(pa, "catAx") == XLSX._slot(pa, "valAx") == XLSX._slot(pa, "dateAx")
        @test XLSX._slot(pa, "layout") < XLSX._slot(pa, "barChart") < XLSX._slot(pa, "catAx") < XLSX._slot(pa, "spPr")
        @test isnothing(XLSX._slot(pa, "nonsense"))
    end

    @testset "insert_child" begin
        sp3 = XLSX.first_element_with_tag(XLSX._series(c, 3).raw, "spPr")
        sp3b = XLSX.insert_child(sp3, SP, mk(XLSX.NS_A, "noFill"))
        @test XLSX.localname.(collect(XML.eachelement(sp3b))) == ["noFill", "ln", "effectLst"]
        @test XLSX.localname.(collect(XML.eachelement(sp3))) == ["ln", "effectLst"]

        @test occursin("noFill", XML.write(XLSX.insert_child(mk(XLSX.NS_C, "spPr"), SP, mk(XLSX.NS_A, "noFill"))))

        # A PARSED childless element has `children === nothing`, unlike a constructed
        # one. This is the case that makes insert_child return a node rather than mutate.
        parsed = XML.parse("""<c:spPr xmlns:c="$(XLSX.NS_C)"/>""", XML.Node)[1]
        @test isnothing(parsed.children)
        @test occursin("noFill", XML.write(XLSX.insert_child(parsed, SP, mk(XLSX.NS_A, "noFill"))))

        bare = XML.parse("""<c:ser xmlns:c="$(XLSX.NS_C)"><c:idx val="0"/><c:order val="0"/><c:val/></c:ser>""", XML.Node)[1]
        got = XLSX.insert_child(bare, (XLSX.NS_C, "barSer"), mk(XLSX.NS_C, "spPr"))
        @test XLSX.localname.(collect(XML.eachelement(got))) == ["idx", "order", "spPr", "val"]

        twice = XLSX.insert_child(got, (XLSX.NS_C, "barSer"), mk(XLSX.NS_C, "spPr"))
        @test count(k -> XLSX.localname(k) == "spPr", XML.children(twice)) == 1

        @test_throws XLSX.XLSXError XLSX.insert_child(bare, (XLSX.NS_C, "nonesuch"), mk(XLSX.NS_C, "spPr"))
        @test_throws XLSX.XLSXError XLSX.insert_child(bare, (XLSX.NS_C, "barSer"), mk(XLSX.NS_C, "axId"))
    end

    @testset "xsd:choice" begin
        ax = XLSX.getChartAxis(c, 1773317264).raw          # has crosses="max"
        @test_throws XLSX.XLSXError XLSX.insert_child(ax, (XLSX.NS_C, "valAx"), mk(XLSX.NS_C, "crossesAt"; val="0"))

        sp = XLSX.insert_child(XLSX.first_element_with_tag(XLSX._series(c, 3).raw, "spPr"), SP, mk(XLSX.NS_A, "noFill"))
        @test_throws XLSX.XLSXError XLSX.insert_child(sp, SP, mk(XLSX.NS_A, "solidFill"))
        # A member does not exclude itself: replacing a fill with the same kind is fine.
        @test !isnothing(XLSX.insert_child(sp, SP, mk(XLSX.NS_A, "noFill")))
    end

    @testset "rebuild_path" begin
        ser3 = XLSX._series(c, 3).raw
        newroot = XLSX.rebuild_path(root,
            [(XLSX.NS_C, "chart") => "chart",
                (XLSX.NS_C, "plotArea") => "plotArea",
                (XLSX.NS_C, "lineChart") => "lineChart",
                (XLSX.NS_C, "lineSer") => ("ser", n -> n === ser3),
                SP => "spPr"],
            sp -> XLSX.insert_child(sp, SP, mk(XLSX.NS_A, "noFill"));
            prefixes=pfx)

        @test newroot !== root
        @test XLSX.first_element_with_tag(newroot, "chart") !== XLSX.first_element_with_tag(root, "chart")
        lc = XLSX.first_element_with_tag(XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(newroot, "chart"), "plotArea"), "lineChart")
        @test occursin("noFill", XML.write(XLSX.first_element_with_tag(
            XLSX.first_element_with_tag(lc, "ser"), "spPr")))

        # A predicate that matches nothing throws rather than creating.
        @test_throws XLSX.XLSXError XLSX.rebuild_path(root,
            [(XLSX.NS_C, "chart") => "chart",
                (XLSX.NS_C, "plotArea") => "plotArea",
                (XLSX.NS_C, "lineChart") => ("lineChart", n -> false)],
            identity; prefixes=pfx)
    end

    @testset "remove_child" begin
        bare = XML.parse("""<c:ser xmlns:c="$(XLSX.NS_C)"><c:idx val="0"/><c:order val="0"/><c:val/></c:ser>""", XML.Node)[1]

        got = XLSX.remove_child(bare, "order")
        @test XLSX.localname.(collect(XML.eachelement(got))) == ["idx", "val"]
        @test XLSX.localname.(collect(XML.eachelement(bare))) == ["idx", "order", "val"]  # original untouched

        # absent child is a no-op, and returns the same object
        @test XLSX.remove_child(bare, "spPr") === bare

        # childless parent is a no-op
        parsed = XML.parse("""<c:spPr xmlns:c="$(XLSX.NS_C)"/>""", XML.Node)[1]
        @test XLSX.remove_child(parsed, "ln") === parsed

        # only the first match goes
        two = XML.parse("""<c:ser xmlns:c="$(XLSX.NS_C)"><c:dPt/><c:dPt/></c:ser>""", XML.Node)[1]
        @test length(XML.children(XLSX.remove_child(two, "dPt"))) == 1
    end

    @testset "remove_choice" begin
        SP = (XLSX.NS_A, "spPr")

        # removes whichever member is present, not the one named
        sp = XLSX.insert_child(XML.parse("""<c:spPr xmlns:c="$(XLSX.NS_C)"/>""", XML.Node)[1],
            SP, mk(XLSX.NS_A, "noFill"))
        @test XLSX.localname.(collect(XML.eachelement(sp))) == ["noFill"]
        @test isempty(XML.children(XLSX.remove_choice(sp, SP, "solidFill")))

        # leaves other slots alone: a line survives fill removal
        sp2 = XLSX.insert_child(XLSX.insert_child(XML.parse("""<c:spPr xmlns:c="$(XLSX.NS_C)"/>""", XML.Node)[1],
            SP, mk(XLSX.NS_A, "noFill")),
            SP, mk(XLSX.NS_A, "ln"))
        @test XLSX.localname.(collect(XML.eachelement(XLSX.remove_choice(sp2, SP, "solidFill")))) == ["ln"]

        # nothing present is a no-op
        lnonly = XLSX.insert_child(XML.parse("""<c:spPr xmlns:c="$(XLSX.NS_C)"/>""", XML.Node)[1],
            SP, mk(XLSX.NS_A, "ln"))
        @test XLSX.localname.(collect(XML.eachelement(XLSX.remove_choice(lnonly, SP, "noFill")))) == ["ln"]

        # a tag that is not in any choice group throws
        @test_throws XLSX.XLSXError XLSX.remove_choice(sp, SP, "ln")
        @test_throws XLSX.XLSXError XLSX.remove_choice(sp, SP, "nonsense")
    end

    @testset "cx CHILD_ORDER agrees with Excel output" begin
        slot(order, name) = findfirst(s -> s isa String ? s == name : name in s, order)
        prefix(tag) = (i=findfirst(':', tag); isnothing(i) ? "" : tag[1:(i-1)])

        function check!(problems, checked, node, nsmap, part)
            ns  = get(nsmap, prefix(XLSX.XML.tag(node)), "")
            ln  = String(XLSX.localname(node))
            key = (ns, ln)
            if !haskey(XLSX.CHILD_ORDER, key) && ln in ("spPr", "txPr", "rich")
                key = (XLSX.NS_A, ln)
            end
            order = get(XLSX.CHILD_ORDER, key, nothing)
            last = 0
            for ch in XLSX.XML.eachelement(node)
                if !isnothing(order) && get(nsmap, prefix(XLSX.XML.tag(ch)), "") in (XLSX.NS_CX, XLSX.NS_A)
                    checked[] += 1
                    k = slot(order, String(XLSX.localname(ch)))
                    if isnothing(k)
                        push!(problems, "$part: $(XLSX.XML.tag(ch)) not in table for $key")
                    elseif k < last
                        push!(problems, "$part: $(XLSX.XML.tag(ch)) out of order in $key")
                    else
                        last = k
                    end
                end
                check!(problems, checked, ch, nsmap, part)
            end
            if !isnothing(order)
                names = [String(XLSX.localname(ch)) for ch in XLSX.XML.eachelement(node)
                     if get(nsmap, prefix(XLSX.XML.tag(ch)), "") in (XLSX.NS_CX, XLSX.NS_A)]
                reps = get(XLSX.REPEATABLE, key, Set{String}())
                for n in unique(names)
                    count(==(n), names) > 1 && !(n in reps) &&
                        push!(problems, "$part: $n repeats in $key but is not in REPEATABLE")
                end
            end
        end

        for fixture in ("chart_ex.xlsx", "chartex_layouts.xlsx", "chartex_formatted.xlsx")
            xf = XLSX.readxlsx(joinpath(data_directory, fixture))
            for c in filter(x -> x isa XLSX.ChartEx, XLSX.getCharts(xf))
                root = XLSX.xml_root_element(xf.data[c.path])
                nsmap = XLSX.get_namespaces(root)
                @test get(nsmap, prefix(XLSX.XML.tag(root)), "") == XLSX.NS_CX   # the root is cx:
                problems = String[]
                checked = Ref(0)
                check!(problems, checked, root, nsmap, "$fixture:$(c.path)")
                @test checked[] > 0                                              # not vacuous
                @test isempty(problems)
                isempty(problems) || foreach(println, problems)
            end
        end
    end

end # ChartSchema