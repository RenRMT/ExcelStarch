Attribute VB_Name = "modExport"
Option Explicit

Public Sub RunChartExport()
#If Mac Then
    MsgBox "Chart export is not supported on Mac.", vbExclamation
    Exit Sub
#End If

    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim sPrompt As String: sPrompt = "Browse to a folder and enter a file name"
    Dim sPathName As String: sPathName = ActiveWorkbook.Path

    If InStr(sPathName, "/") Or Len(sPathName) = 0 Then
        sPathName = CurDir
    End If

    Dim sFileExt As String
    sFileExt = GetSetting(exportAppName, exportSection, exportSettingKey, exportDefaultExt)

    Dim sFilters As String
    sFilters = "PNG Files (*.png),*.png," & _
               "GIF Files (*.gif),*.gif," & _
               "JPEG Files (*.jpeg;*.jpe;*.jpg),*.jpeg;*.jpe;*.jpg," & _
               "BMP Files (*.bmp),*.bmp," & _
               "SVG Files (*.svg),*.svg," & _
               "PDF Files (*.pdf),*.pdf"

    Dim vFileExt
    vFileExt = Array("*", "png", "gif", "jpg", "bmp", "svg", "pdf")

    Dim iFilterIndex As Long
    On Error Resume Next
    iFilterIndex = WorksheetFunction.Match(sFileExt, vFileExt, 0) - 1
    On Error GoTo 0
    If iFilterIndex = 0 Then iFilterIndex = 1: sFileExt = exportDefaultExt

    Dim sFileName As String
    sFileName = sPathName & "\" & exportDefaultName & "." & sFileExt

    Dim vChartName As Variant
    vChartName = Application.GetSaveAsFilename(InitialFileName:=sFileName, _
                                               FileFilter:=sFilters, _
                                               FilterIndex:=iFilterIndex, _
                                               Title:=sPrompt)
    If VarType(vChartName) = vbBoolean Then Exit Sub

    sFileName = vChartName
    sFileExt = Mid$(sFileName, InStrRev(sFileName, ".") + 1)

    Dim FileFilter As String, bPDF As Boolean

    Select Case LCase$(sFileExt)
        Case "png": FileFilter = "PNG"
        Case "jpeg", "jpe", "jpg": FileFilter = "JPG"
        Case "bmp": FileFilter = "BMP"
        Case "gif": FileFilter = "GIF"
        Case "svg": FileFilter = "SVG"
        Case "pdf": bPDF = True
        Case Else
            sFileExt = exportDefaultExt
            sFileName = sFileName & "." & exportDefaultExt
            FileFilter = "PNG"
    End Select

    If Not bPDF Then
        ActiveChart.Export sFileName, FileFilter
    Else
        Select Case TypeName(ActiveChart.Parent)
            Case "ChartObject"
                ActiveChart.Parent.Parent.Activate
                ActiveChart.ChartArea.Select
            Case "Workbook"
                ActiveChart.Select
        End Select

        DoEvents

        ActiveChart.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sFileName, _
                                        Quality:=xlQualityStandard, _
                                        IncludeDocProperties:=False, _
                                        IgnorePrintAreas:=False, _
                                        OpenAfterPublish:=False
    End If

    SaveSetting exportAppName, exportSection, exportSettingKey, sFileExt
End Sub

Sub ExportChart()
    RunChartExport
End Sub

Public Sub ChartExport_onAction(control As IRibbonControl)
    RunChartExport
End Sub
