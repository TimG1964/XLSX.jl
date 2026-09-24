# Charts

```@meta
CurrentModule = XLSX.Charts
```

Chart support lives in the `XLSX.Charts` sub-module. `using XLSX.Charts` brings
the functions below into scope; the handle and value types are public but not
exported, so refer to them as `XLSX.Charts.ChartSeries` and so on. They are 
documented on the [Chart types](chartTypes.md) page.

The names released before the sub-module existed — `getCharts`, `getChart`,
`getChartData`, `getChartRanges`, `chartType`, `chartSchema`, `AbstractChart`,
`Chart`, `ChartEx`, `ChartRef` and `ChartSeries` — are also reachable as
`XLSX.getCharts` and so on.

```@docs
XLSX.Charts
```

## Finding and reading charts

```@docs
getCharts
getChart
getChartSchema
getChartType
getChartTitle
getChartData
getChartRanges
```

## The chart as a whole

```@docs
getChartTypes
getChartGroups
getChartLegend
getLegendPos
getLegendOverlay
getLegendShapeProps
getLegendTextProps
setLegendTextProp
getChartTitleText
getChartTitleRef
getChartTitleShapeProps
getChartTitleTextProps
getAutoTitleDeleted
setChartTitleText
setChartTitleTextProp
getPlotAreaShapeProps
getChartSpaceShapeProps
getChartSpaceTextProps
setChartSpaceTextProp
```

## Series

```@docs
getChartSeries
getSeriesGroup
getSeriesAxes
getSeriesShapeProps
getSeriesFill
getSeriesLine
getSeriesMarker
getSeriesLabelTextProps
setSeriesFill
setSeriesLine
setSeriesLineColor
setSeriesLineWidth
setSeriesLineDash
setSeriesLineCap
setSeriesLineCompound
setSeriesLineJoin
setSeriesLineMiterLimit
has_fill
has_line
```

## Chart groups

```@docs
getGroupAxes
getGroupLabelTextProps
getGroupDropLines
getGroupHiLowLines
getGroupSeriesLines
getGroupUpDownBars
getUpBarShapeProps
getDownBarShapeProps
setGroupLabelTextProp
```

## Axes

```@docs
getChartAxes
getChartAxis
getAxisPartner
getAxisShapeProps
getAxisTextProps
getAxisGridlines
getAxisTitleText
getAxisTitleRef
getAxisTitleTextProps
setAxisTitleText
setAxisTitleTextProp
getAxisNumberFormatCode
getAxisNumberFormatLinked
getAxisMajorTickMark
getAxisMinorTickMark
getAxisTickLabelPos
getAxisOrientation
getAxisMin
getAxisMax
getAxisLogBase
getAxisCrosses
getAxisCrossesAt
getAxisCrossBetween
getAxisMajorUnit
getAxisMinorUnit
getAxisLabelAlign
getAxisLabelOffset
getAxisMultiLevelLabels
```

## Data points, markers and labels

```@docs
getSeriesDataPoints
getSeriesDataPoint
getDataPointShapeProps
getDataPointMarker
getMarkerFill
setMarker
setMarkerSymbol
setMarkerSize
setMarkerFill
setMarkerLineColor
setMarkerLineWidth
getSeriesDataLabels
getSeriesDataLabel
getDataLabelText
getDataLabelTextProps
getDataLabelShapeProps
getDataLabelPosition
getDataLabelOffset
getLabelTextProp
setLabelText
setLabelTextProp
setLabelDeleted
```

## Trendlines and error bars

```@docs
getSeriesTrendlines
getTrendlineShapeProps
getTrendlineLabelText
getTrendlineLabelTextProps
getTrendlineLabelShapeProps
getSeriesErrorBars
getErrorBarsShapeProps
getErrorBarsCustomRefs
```

## Creating charts

```@docs
addChart
addChartEx
addSeries
```
