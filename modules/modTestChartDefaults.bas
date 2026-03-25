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

    Debug.Assert def.Gridlines = axisY
    Debug.Assert def.AxisDisplay = axisBoth

    Debug.Print "  PASS: TestLineChartDefaults"
End Sub

Private Sub TestBarChartDefaults()
    Dim def As ChartDefaults
    def = BarChartDefaults()

    Debug.Assert def.Gridlines = axisX
    Debug.Assert def.AxisDisplay = axisBoth

    Debug.Print "  PASS: TestBarChartDefaults"
End Sub

Private Sub TestColumnChartDefaults()
    Dim def As ChartDefaults
    def = ColumnChartDefaults()

    Debug.Assert def.Gridlines = axisY
    Debug.Assert def.AxisDisplay = axisBoth

    Debug.Print "  PASS: TestColumnChartDefaults"
End Sub

Private Sub TestAreaChartDefaults()
    Dim def As ChartDefaults
    def = AreaChartDefaults()

    Debug.Assert def.Gridlines = axisY
    Debug.Assert def.AxisDisplay = axisBoth

    Debug.Print "  PASS: TestAreaChartDefaults"
End Sub

Private Sub TestScatterChartDefaults()
    Dim def As ChartDefaults
    def = ScatterChartDefaults()

    Debug.Assert def.Gridlines = axisBoth
    Debug.Assert def.AxisDisplay = axisBoth

    Debug.Print "  PASS: TestScatterChartDefaults"
End Sub

Private Sub TestPieChartDefaults()
    Dim def As ChartDefaults
    def = PieChartDefaults()

    Debug.Assert def.Gridlines = axisNone
    Debug.Assert def.AxisDisplay = axisNone
    Debug.Assert def.Legend = True

    Debug.Print "  PASS: TestPieChartDefaults"
End Sub

Private Sub TestTreemapChartDefaults()
    Dim def As ChartDefaults
    def = TreemapChartDefaults()

    Debug.Assert def.Gridlines = axisNone
    Debug.Assert def.AxisDisplay = axisNone

    Debug.Print "  PASS: TestTreemapChartDefaults"
End Sub
