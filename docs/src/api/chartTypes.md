# Chart types

```@meta
CurrentModule = XLSX.Charts
```

The types the chart API takes and returns: handles, the value types read from a
chart, and the DrawingML structs that carry formatting. The functions are on the
[Charts](charts.md) page.

## chartEx charts

Charts in the Microsoft `cx:` namespace — waterfall, funnel, treemap, sunburst,
histogram, Pareto, box & whisker and region map. Many of the functions above
also take a [`ChartEx`](@ref); these are specific to it.

```@docs
getChartSeriesCount
getChartDataBlocks
getChartAxisIds
getChartTitleRange
getSeriesLayout
getSeriesLayoutFlag
getSeriesData
getSeriesName
getSeriesNameRange
getSeriesOwner
getSeriesHidden
getSeriesAxisIds
getSeriesSubtotals
getSeriesBinning
getSeriesAggregation
getSeriesQuartileMethod
getSeriesParentLabelLayout
getLabelFlag
getLabelPosition
setSeriesName
setSeriesLayoutFlag
setSeriesSubtotals
setSeriesBinning
setSeriesAggregation
setSeriesQuartileMethod
setSeriesParentLabelLayout
```

## Handles and values

```@docs
AbstractChart
Chart
ChartEx
ChartRef
ChartRange
ChartRanges
ChartSeries
ChartAxis
ChartGroup
ChartMarker
ChartDataPoint
ChartDataLabel
ChartTrendline
ChartErrorBars
ChartUpDownBars
ChartExData
ChartExDimension
ChartExBinning
FormatSite
Effective
```

## Colours and DrawingML

```@docs
SchemeColor
ColorSpec
DrawingColor
DrawingFill
DrawingLine
DrawingShapeProps
DrawingText
DrawingParagraph
DrawingParaProps
DrawingRun
DrawingRunProps
DrawingBodyProps
```