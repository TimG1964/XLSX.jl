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
end
