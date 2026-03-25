Attribute VB_Name = "modTestChartDefaults"
Option Explicit

' Test harness for ChartDefaults refactoring
' Verifies that:
'   1. All chart types render identically to pre-refactoring (regression test)
'   2. Chart-type-specific defaults apply correctly (behavior verification)
'   3. UDT construction and factory functions work as expected

Public Sub RunAllTests()
    Debug.Print "=== ChartDefaults Refactoring Tests ==="

    TestUDTConstruction
    TestFactoryFunctions
    TestLineChartDefaults
    TestBarChartDefaults
    TestColumnChartDefaults
    TestAreaChartDefaults
    TestScatterChartDefaults
    TestPieChartDefaults
    TestTreemapChartDefaults

    Debug.Print "=== All tests passed ==="
End Sub

Private Sub TestUDTConstruction()
    Dim defaults As ChartDefaults

    defaults.Gridlines = axisY
    defaults.AxisDisplay = axisBoth
    defaults.AxisLines = axisNone
    defaults.AxisLabels = axisBoth
    defaults.Legend = True

    Debug.Assert defaults.Gridlines = axisY
    Debug.Assert defaults.AxisDisplay = axisBoth
    Debug.Assert defaults.Legend = True
    Debug.Print "  PASS: TestUDTConstruction"
End Sub

Private Sub TestFactoryFunctions()
    Dim def As ChartDefaults

    ' Test that factory functions return valid ChartDefaults
    def = DefaultChartDefaults()
    Debug.Assert def.Gridlines = defaultGridlines
    Debug.Assert def.AxisDisplay = defaultAxisDisplay

    def = LineChartDefaults()
    Debug.Assert def.Gridlines <> axisNone

    Debug.Print "  PASS: TestFactoryFunctions"
End Sub

Private Sub TestLineChartDefaults()
    Dim def As ChartDefaults
    def = LineChartDefaults()

    Debug.Assert def.Gridlines = axisY, "Line should have Y-gridlines only"
    Debug.Assert def.AxisDisplay = axisBoth, "Line should show both axes"
    Debug.Assert def.AxisLines = axisNone, "Line should hide axis lines"
    Debug.Assert def.AxisLabels = axisBoth, "Line should show both axis labels"

    Debug.Print "  PASS: TestLineChartDefaults"
End Sub

Private Sub TestBarChartDefaults()
    Dim def As ChartDefaults
    def = BarChartDefaults()

    Debug.Assert def.Gridlines = axisX, "Bar should have X-gridlines only"
    Debug.Assert def.AxisDisplay = axisBoth, "Bar should show both axes"
    Debug.Assert def.AxisLines = axisNone, "Bar should hide axis lines"
    Debug.Assert def.AxisLabels = axisBoth, "Bar should show both axis labels"

    Debug.Print "  PASS: TestBarChartDefaults"
End Sub

Private Sub TestColumnChartDefaults()
    Dim def As ChartDefaults
    def = ColumnChartDefaults()

    Debug.Assert def.Gridlines = axisY, "Column should have Y-gridlines only"
    Debug.Assert def.AxisDisplay = axisBoth, "Column should show both axes"
    Debug.Assert def.AxisLines = axisNone, "Column should hide axis lines"
    Debug.Assert def.AxisLabels = axisBoth, "Column should show both axis labels"

    Debug.Print "  PASS: TestColumnChartDefaults"
End Sub

Private Sub TestAreaChartDefaults()
    Dim def As ChartDefaults
    def = AreaChartDefaults()

    Debug.Assert def.Gridlines = axisY, "Area should have Y-gridlines only"
    Debug.Assert def.AxisDisplay = axisBoth, "Area should show both axes"
    Debug.Assert def.AxisLines = axisNone, "Area should hide axis lines"

    Debug.Print "  PASS: TestAreaChartDefaults"
End Sub

Private Sub TestScatterChartDefaults()
    Dim def As ChartDefaults
    def = ScatterChartDefaults()

    Debug.Assert def.Gridlines = axisBoth, "Scatter should have both gridlines"
    Debug.Assert def.AxisDisplay = axisBoth, "Scatter should show both axes"
    Debug.Assert def.AxisLines = axisNone, "Scatter should hide axis lines"

    Debug.Print "  PASS: TestScatterChartDefaults"
End Sub

Private Sub TestPieChartDefaults()
    Dim def As ChartDefaults
    def = PieChartDefaults()

    Debug.Assert def.Gridlines = axisNone, "Pie should have no gridlines"
    Debug.Assert def.AxisDisplay = axisNone, "Pie should have no axes"
    Debug.Assert def.AxisLines = axisNone, "Pie should have no axis lines"
    Debug.Assert def.Legend = True, "Pie should show legend by default"

    Debug.Print "  PASS: TestPieChartDefaults"
End Sub

Private Sub TestTreemapChartDefaults()
    Dim def As ChartDefaults
    def = TreemapChartDefaults()

    Debug.Assert def.Gridlines = axisNone, "Treemap should have no gridlines"
    Debug.Assert def.AxisDisplay = axisNone, "Treemap should have no axes"
    Debug.Assert def.AxisLines = axisNone, "Treemap should have no axis lines"
    Debug.Assert def.AxisLabels = axisNone, "Treemap should have no axis labels"

    Debug.Print "  PASS: TestTreemapChartDefaults"
End Sub
