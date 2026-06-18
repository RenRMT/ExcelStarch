Attribute VB_Name = "modConfigCharts"

Option Explicit
' +---------------------------------------------------------+
' |  DEFAULT CHART FORMATTING                               |
' |  Controls pipeline defaults for new charts.             |
' |  Axis constants: 0=none, 1=X only, 2=Y only, 3=both    |
' +---------------------------------------------------------+
'=== Axis selection values (used by defaultGridlines, defaultAxisDisplay, etc.) ===
Public Const axisNone As Long = 0
Public Const axisX As Long = 1
Public Const axisY As Long = 2
Public Const axisBoth As Long = 3

'=== Data label contrast settings ===
Public Const wcagLuminanceThreshold As Double = 0.179   ' WCAG threshold for black vs white label text

'=== Scatter / bubble styling ===
Public Const scatterMarkerSize As Long = 9              ' marker point size (points) for scatter charts
Public Const bubbleTransparency As Single = 0.5         ' bubble fill transparency (0 = opaque, 1 = clear)

'=== Default formatting for new/reformatted charts ===
Public Const defaultGridlines As Long = axisNone        ' gridline visibility
Public Const defaultAxisDisplay As Long = axisNone      ' axis visibility (HasAxis)
Public Const defaultAxisLines As Long = axisNone        ' axis line visibility
Public Const defaultAxisLabels As Long = axisNone       ' tick label visibility
Public Const defaultLegend As Boolean = False           ' False = no legend

'=== ChartDefaults User-Defined Type ===
'Bundles formatting options into a single parameter for chart pipeline.
'Only Gridlines, AxisDisplay, and Legend are currently consumed by ApplyDefaultFormatting.
'ShowYAxisTitle is consumed only by the chartex chrome (modChartExChrome) - it adds the
'optional worksheet Y-axis title box for chartex types with a value axis (box & whisker);
'classic factories leave it False and ignore it.
'AxisLines and AxisLabels are reserved for future use (phase 6+).
Public Type ChartDefaults
    Gridlines As Long       ' axisNone, axisX, axisY, axisBoth (controls gridline visibility)
    AxisDisplay As Long     ' axisNone, axisX, axisY, axisBoth (controls axis visibility)
    Legend As Boolean       ' True = show legend, False = hide
    ShowYAxisTitle As Boolean  ' chartex only: add a worksheet Y-axis title box
End Type

'=== Factory function for global defaults ===
Public Function DefaultChartDefaults() As ChartDefaults
    With DefaultChartDefaults
        .Gridlines = defaultGridlines
        .AxisDisplay = defaultAxisDisplay
        .Legend = defaultLegend
    End With
End Function

'=== Chart-type-specific profile factories ===
Public Function LineChartDefaults() As ChartDefaults
    With LineChartDefaults
        .Gridlines = axisY          ' Y-gridlines only (horizontal lines showing value scale)
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function BarChartDefaults() As ChartDefaults
    With BarChartDefaults
        .Gridlines = axisX          ' X-gridlines only (vertical lines showing value scale)
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function ColumnChartDefaults() As ChartDefaults
    With ColumnChartDefaults
        .Gridlines = axisY          ' Y-gridlines only (horizontal lines showing value scale)
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function AreaChartDefaults() As ChartDefaults
    With AreaChartDefaults
        .Gridlines = axisY          ' Y-gridlines only (horizontal lines showing value scale)
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function ScatterChartDefaults() As ChartDefaults
    With ScatterChartDefaults
        .Gridlines = axisBoth       ' Both gridlines for reference grid
        .AxisDisplay = axisBoth     ' Show both X and Y axes
        .Legend = defaultLegend     ' Use global default
    End With
End Function

Public Function PieChartDefaults() As ChartDefaults
    With PieChartDefaults
        .Gridlines = axisNone       ' No gridlines (pie has no axes)
        .AxisDisplay = axisNone     ' No axes for pie charts
        .Legend = True              ' Pie typically shows legend for slice labels
    End With
End Function

Public Function TreemapChartDefaults() As ChartDefaults
    With TreemapChartDefaults
        .Gridlines = axisNone       ' No gridlines (treemap has no axes)
        .AxisDisplay = axisNone     ' No axes for treemaps
        .Legend = defaultLegend     ' Use global default (tile labels usually suffice)
        .ShowYAxisTitle = False     ' Treemap has no value axis - no Y-axis title box
    End With
End Function
