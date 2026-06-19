Attribute VB_Name = "modExport"

Option Explicit

Public Sub RunChartExport()
    ' Mac Excel lacks GetSaveAsFilename and Chart.Export in the form used below,
    ' so export is Windows-only; bail out with an explanatory message on Mac.
#If Mac Then
    MsgExportMacUnsupported
    Exit Sub
#End If

    On Error GoTo CleanFail

    'Chartex chrome lives in a worksheet group (not inside the chart), so Chart.Export
    'would omit it. If a chartex group is selected, rasterise the whole group instead.
    'Returns False when there is no chartex group, falling through to the classic path.
    If TrySaveChartExGroupAsPicture() Then Exit Sub

    'Ensure a chart is active
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    'Explicit chart reference avoids relying on ActiveChart later
    Dim targetChart As Chart
    Set targetChart = ActiveChart

    '----- File dialog setup -----
    Dim Prompt As String: Prompt = "Browse to a folder and enter a file name"
    Dim PathName As String: PathName = ActiveWorkbook.Path

    If InStr(PathName, "/") > 0 Or Len(PathName) = 0 Then
        PathName = CurDir
    End If

    Dim FileExt As String
    FileExt = GetSetting(exportAppName, exportSection, exportSettingKey, exportDefaultExt)

    Dim Filters As String
    Filters = "PNG Files (*.png),*.png," & _
              "GIF Files (*.gif),*.gif," & _
              "JPEG Files (*.jpeg;*.jpe;*.jpg),*.jpeg;*.jpe;*.jpg," & _
              "BMP Files (*.bmp),*.bmp," & _
              "SVG Files (*.svg),*.svg," & _
              "PDF Files (*.pdf),*.pdf"

    Dim FileExtArray As Variant
    FileExtArray = Array("*", "png", "gif", "jpg", "bmp", "svg", "pdf")

    ' Match raises an error (rather than returning an error value) when the saved
    ' FileExt is not in the list; suppress it so an unrecognised setting falls
    ' through to the default below instead of aborting the export.
    Dim FilterIndex As Long
    On Error Resume Next
    FilterIndex = WorksheetFunction.Match(FileExt, FileExtArray, 0) - 1
    On Error GoTo 0
    If FilterIndex = 0 Then
        FilterIndex = 1
        FileExt = exportDefaultExt
    End If

    Dim FileName As String
    FileName = PathName & "\" & exportDefaultName & "." & FileExt

    Dim dialogResult As Variant
    dialogResult = Application.GetSaveAsFilename( _
        InitialFileName:=FileName, _
        FileFilter:=Filters, _
        FilterIndex:=FilterIndex, _
        Title:=Prompt _
    )

    If VarType(dialogResult) = vbBoolean Then Exit Sub 'User canceled

    FileName = dialogResult
    FileExt = Mid$(FileName, InStrRev(FileName, ".") + 1)

    'Normalize JPEG variants
    Select Case LCase$(FileExt)
        Case "jpeg", "jpe": FileExt = "jpg"
    End Select

    '----- Determine export format -----
    Dim FileFilter As String
    Dim isPDF As Boolean

    Select Case LCase$(FileExt)
        Case "png": FileFilter = "PNG"
        Case "jpg": FileFilter = "JPG"
        Case "bmp": FileFilter = "BMP"
        Case "gif": FileFilter = "GIF"
        Case "svg": FileFilter = "SVG"
        Case "pdf": isPDF = True
        Case Else
            FileExt = exportDefaultExt
            FileName = FileName & "." & exportDefaultExt
            FileFilter = "PNG"
    End Select

    '----- Perform export -----
    If Not isPDF Then
        'Use explicit chart reference (no ActiveChart)
        targetChart.Export FileName, FileFilter

    Else
        'PDF export requires that the exporting chart or its sheet be active
        Dim parentObj As Object
        Set parentObj = targetChart.Parent

        Select Case TypeName(parentObj)
            Case "ChartObject"
                parentObj.Parent.Activate     'Activate worksheet
                targetChart.Activate          'Ensure the chart is active

            Case "Workbook"
                targetChart.Activate          'Chart sheet
        End Select

        DoEvents 'Allow Excel to prepare chart render pipeline

        targetChart.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            FileName:=FileName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    End If

    'Save chosen extension as user preference
    SaveSetting exportAppName, exportSection, exportSettingKey, FileExt
    Exit Sub

CleanFail:
    MsgError "RunChartExport"
End Sub

Sub ExportChart()
    RunChartExport
End Sub


' Chartex chrome (title/logo/source) is built as worksheet shapes grouped with the
' chart (see modEngineExChrome), so the chart-only Chart.Export omits it. This exports
' the whole group as a PNG by rasterising it through a temporary chart.
'
' Returns True if a chartex group was found and handled (or the user cancelled the
' dialog); returns False when no chartex group is selected, so RunChartExport falls
' through to the classic Chart.Export path.
'
' Note: the group is rasterised at ~screen resolution (CopyPicture xlScreen) - softer
' than Chart.Export. PNG only, by design (see plan).
Private Function TrySaveChartExGroupAsPicture() As Boolean
#If Mac Then
    'Export is Windows-only (see RunChartExport); leave the classic path to message.
    TrySaveChartExGroupAsPicture = False
    Exit Function
#End If

    On Error GoTo CleanFail

    Dim grp As Shape
    Set grp = ResolveChartExGroup()
    If grp Is Nothing Then
        TrySaveChartExGroupAsPicture = False
        Exit Function
    End If

    'From here we own the request: always return True so the classic path is skipped.
    TrySaveChartExGroupAsPicture = True

    'File dialog (PNG only for the chartex group path).
    Dim PathName As String: PathName = ActiveWorkbook.Path
    If InStr(PathName, "/") > 0 Or Len(PathName) = 0 Then PathName = CurDir

    Dim FileName As String
    FileName = PathName & "\" & exportDefaultName & ".png"

    Dim dialogResult As Variant
    dialogResult = Application.GetSaveAsFilename( _
        InitialFileName:=FileName, _
        FileFilter:="PNG Files (*.png),*.png", _
        Title:="Browse to a folder and enter a file name")

    If VarType(dialogResult) = vbBoolean Then Exit Function 'User cancelled
    FileName = dialogResult

    SaveChartExGroupPng grp, FileName

    Exit Function

CleanFail:
    'Already flagged True if we owned the request; report and stop.
    MsgError "TrySaveChartExGroupAsPicture"
End Function


' Resolves the chartex group from the current selection: the selected shape may be
' the group itself, or a member whose parent group is a chartex group. Returns
' Nothing if the selection is not (part of) a chartex group.
Private Function ResolveChartExGroup() As Shape
    Dim shp As Shape

    'ShapeRange(1) raises when the selection is a cell range, not a shape - probe it
    'under a narrow handler rather than a function-wide blanket suppressor.
    On Error Resume Next
    Set shp = Selection.ShapeRange(1)
    On Error GoTo 0
    If shp Is Nothing Then Exit Function

    If shp.Type = msoGroup Then
        If IsChartExGroupName(shp.name) Then Set ResolveChartExGroup = shp
        Exit Function
    End If

    'ParentGroup RAISES (it does not return Nothing) when shp is not in a group, so
    'the suppression here is load-bearing: a raise simply means "not grouped".
    Dim parent As Shape
    On Error Resume Next
    Set parent = shp.ParentGroup
    On Error GoTo 0
    If Not parent Is Nothing Then
        If IsChartExGroupName(parent.name) Then Set ResolveChartExGroup = parent
    End If
End Function


Private Function IsChartExGroupName(ByVal nm As String) As Boolean
    'chartExGroupPrefix is the single source of truth, defined in modEngineExChrome.
    IsChartExGroupName = (InStr(1, nm, chartExGroupPrefix, vbTextCompare) = 1)
End Function


' Rasterises a shape group to PNG via a temporary chart canvas (Excel has no
' Shape.Export). The temp chart is always removed.
Private Sub SaveChartExGroupPng(grp As Shape, ByVal FileName As String)
    On Error GoTo CleanFail
    AppFast   'suppress the flash of the temporary ChartObject

    Dim ws As Worksheet
    Set ws = grp.Parent     'worksheet hosting the group

    grp.CopyPicture Appearance:=xlScreen, Format:=xlBitmap

    Dim chtObj As ChartObject
    Set chtObj = ws.ChartObjects.Add(Left:=0, Top:=0, Width:=grp.Width, Height:=grp.Height)
    chtObj.Chart.Paste
    chtObj.Chart.Export FileName, "PNG"

    chtObj.Delete

    'Save chosen extension as user preference (mirrors the classic path).
    SaveSetting exportAppName, exportSection, exportSettingKey, "png"
    AppRestore
    Exit Sub

CleanFail:
    'Best-effort cleanup of the temp chart before surfacing the error.
    On Error Resume Next
    If Not chtObj Is Nothing Then chtObj.Delete
    On Error GoTo 0
    AppRestore
    MsgError "SaveChartExGroupPng"
End Sub
