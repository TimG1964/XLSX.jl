using Documenter, XLSX
makedocs(
    sitename = "XLSX.jl",
    modules = [XLSX, XLSX.Charts],
    pages = [
        "Home" => "index.md",
        "Tutorial" => Any[
            "XLSX Tutorial" => "tutorial/XLSXtutorial.md",
            "FileIO Tutorial" => "tutorial/FileIOtutorial.md",
        ],
        "Formatting Guide" => Any[
            "Cell formats" => "formatting/cellFormatting.md",
            "Conditional formats" => "formatting/conditionalFormatting.md",
            "Column width and row height" => "formatting/widthAndHeight.md",
            "Merged cells" => "formatting/mergedCells.md",
            "Freeze/split panes" => "formatting/freezeAndSplitPanes.md",
        ],
        "Using Formulas" => "formulae/formulas.md",
        "Using Excel Tables" => "tables/excelTables.md",
        "Excel Charts" => Any[
            "Overview" => "charts/excelCharts.md",
            "Reading charts" => "charts/readingCharts.md",
            "Formatting charts" => "charts/formattingCharts.md",
            "Creating charts" => "charts/creatingCharts.md",
            "Creating chartEx charts" => "charts/creatingChartEx.md",
            "Sheets with charts" => "charts/chartSheets.md",
            "Limitations" => "charts/chartLimitations.md",
        ],
        "Examples" => "examples.md",
        "Migration Guide" => "migration.md",
        "API Reference" => Any[
            "Files and worksheets" => "api/files.md",
            "Cells and data" => "api/data.md",
            "Formats" => "api/formats.md",
            "Charts" => "api/charts.md",
            "Chart types" => "api/chartTypes.md",
        ]
    ],
    checkdocs = :public,
)
deploydocs(
    repo = "github.com/JuliaData/XLSX.jl.git",
    target = "build",
    versions = [
        "stable" => "v^",
        "dev" => "dev"
    ],
)
