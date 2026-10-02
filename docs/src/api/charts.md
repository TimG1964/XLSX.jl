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
getLegendEntryDeleted
setLegendEntryDeleted
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
setPlotAreaLine
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
hasFill
hasLine
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
getGroupGapWidth
setGroupGapWidth
getGroupOverlap
setGroupOverlap
```

## Creating charts

```@docs
addChart
addChartEx
addSeries
```

