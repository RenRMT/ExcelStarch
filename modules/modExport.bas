Attribute VB_Name = "modExport"

Option Explicit

Public Sub RunChartExport()
#If Mac Then
    MsgExportMacUnsupported
    Exit Sub
#End If

    On Error GoTo CleanFail

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
