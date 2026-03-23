Attribute VB_Name = "modRibbonHandlers"
'==== Module: modRibbonHandlers ====
' All ribbon onAction callbacks
' Each handler is a thin wrapper calling the public entry point in its implementation module.
' Mirrors the ribbon XML structure: one file to audit all button-to-handler mappings.
Option Explicit

'=== Chart creation ===
Public Sub Bar_onAction(control As IRibbonControl): BarChart: End Sub
Public Sub StackedBar_onAction(control As IRibbonControl): StackedBarChart: End Sub
Public Sub Lollipop_onAction(control As IRibbonControl): LollipopChart: End Sub

Public Sub Column_onAction(control As IRibbonControl): ColumnChart: End Sub
Public Sub StackedColumn_onAction(control As IRibbonControl): StackedColumnChart: End Sub

Public Sub LineChart_onAction(control As IRibbonControl): LineChart: End Sub
Public Sub Pie_onAction(control As IRibbonControl): PieChart: End Sub
Public Sub Donut_onAction(control As IRibbonControl): DonutChart: End Sub
Public Sub Area_onAction(control AS IRibbonControl): AreaChart End Sub
'=== Chart tools ===
Public Sub RemoveLegendResizeButton_onAction(control As IRibbonControl): RemoveLegendResizeButton: End Sub
Public Sub LabelLastPointButton_onAction(control As IRibbonControl): LabelLastPointButton: End Sub
Public Sub ToggleGridlinesButton_onAction(control As IRibbonControl): ToggleGridlines: End Sub
Public Sub ToggleAxesButton_onAction(control As IRibbonControl): ToggleAxes: End Sub
Public Sub ToggleDataLabelsButton_onAction(control As IRibbonControl): ToggleDataLabels: End Sub
Public Sub ApplyChartStyleButton_onAction(control As IRibbonControl): ApplyChartStyle: End Sub

'=== Colour ramps ===
Public Sub ApplyRamp_onAction(control As IRibbonControl): ApplyColorRamp control.Tag: End Sub
Public Sub ApplyDivergingRamp_onAction(control As IRibbonControl): ApplyDivergingRampFromTag control.Tag: End Sub
Public Sub InvertRamp_onAction(control As IRibbonControl): InvertColorRamp: End Sub

'=== Colour tools ===
Public Sub TogglePaletteOrder_onAction(control As IRibbonControl): TogglePaletteOrder: End Sub
Public Sub Format_onAction(control As IRibbonControl): ApplyFillFromTag control.Tag: End Sub
Public Sub StartWithGrayButton_onAction(control As IRibbonControl): StartWithGray: End Sub

'=== Utilities ===
Public Sub ChartExport_onAction(control As IRibbonControl): ExportChart: End Sub
