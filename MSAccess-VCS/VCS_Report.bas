Attribute VB_Name = "VCS_Report"
Option Compare Database

Option Private Module
Option Explicit

Private Type str_DEVMODE
    RGB As String * 94
End Type

Private Type type_DEVMODE
    strDeviceName(31) As Byte 'vba strings are encoded in unicode (16 bit) not ascii
    intSpecVersion As Integer
    intDriverVersion As Integer
    intSize As Integer
    intDriverExtra As Integer
    lngFields As Long
    intOrientation As Integer
    intPaperSize As Integer
    intPaperLength As Integer
    intPaperWidth As Integer
    intScale As Integer
    intCopies As Integer
    intDefaultSource As Integer
    intPrintQuality As Integer
    intColor As Integer
    intDuplex As Integer
    intResolution As Integer
    intTTOption As Integer
    intCollate As Integer
    strFormName(31) As Byte
    lngPad As Long
    lngBits As Long
    lngPW As Long
    lngPH As Long
    lngDFI As Long
    lngDFr As Long
End Type

Private Function GetPrtDevModeForReport(ByRef rpt As Report, ByRef vars As type_DEVMODE) As Boolean
    Dim DevModeString As str_DEVMODE
    Dim DevModeExtra As String

    ' Read print vars into struct
    If Not IsNull(rpt.PrtDevMode) Then
        DevModeExtra = rpt.PrtDevMode
        DevModeString.RGB = DevModeExtra
        LSet vars = DevModeString
        GetPrtDevModeForReport = True
    Else
        GetPrtDevModeForReport = False
        Exit Function
    End If
End Function

Private Sub SetPrtDevModeForReport(ByRef rpt As Report, ByRef vars As type_DEVMODE)
    Dim DevModeString As str_DEVMODE
    Dim DevModeExtra As String

    ' Write print vars back into report
    LSet DevModeString = vars
    Mid(DevModeExtra, 1, 94) = DevModeString.RGB
    rpt.PrtDevMode = DevModeExtra
End Sub

' Exports print vars for reports
Public Sub VCS_ExportPrintVars(ByVal obj_name As String, ByVal filePath As String)
    DoEvents

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim DM As type_DEVMODE
    Dim rpt As Report

    ' Report must be open to access Report object
    ' Report must be opened in design view to save changes to the print vars
    DoCmd.OpenReport obj_name, acViewDesign
    Set rpt = Reports(obj_name)

    If GetPrtDevModeForReport(rpt, DM) Then
        Dim OutFile As Object
        Set OutFile = FSO.CreateTextFile(filePath, overwrite:=True, Unicode:=False)

        ' Print out print var values
        OutFile.WriteLine DM.intOrientation
        OutFile.WriteLine DM.intPaperSize
        OutFile.WriteLine DM.intPaperLength
        OutFile.WriteLine DM.intPaperWidth
        OutFile.WriteLine DM.intScale
        OutFile.Close
    Else
        Set rpt = Nothing
        DoCmd.Close acReport, obj_name, acSaveNo
        Debug.Print "Warning: PrtDevMode is null"
        Exit Sub
    End If

    Set rpt = Nothing
    DoCmd.Close acReport, obj_name, acSaveYes
End Sub

Public Sub VCS_ImportPrintVars(ByVal obj_name As String, ByVal filePath As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim DM As type_DEVMODE
    Dim rpt As Report

    ' Report must be open to access Report object
    ' Report must be opened in design view to save changes to the print vars
    DoCmd.OpenReport obj_name, acViewDesign
    Set rpt = Reports(obj_name)

    If GetPrtDevModeForReport(rpt, DM) Then
        Dim InFile As Object
        Set InFile = FSO.OpenTextFile(filePath, iomode:=ForReading, create:=False, Format:=TristateFalse)

        ' Set print var values
        DM.intOrientation = InFile.ReadLine
        DM.intPaperSize = InFile.ReadLine
        DM.intPaperLength = InFile.ReadLine
        DM.intPaperWidth = InFile.ReadLine
        DM.intScale = InFile.ReadLine
        InFile.Close

        SetPrtDevModeForReport rpt, DM
    Else
        Set rpt = Nothing
        DoCmd.Close acReport, obj_name, acSaveNo
        Debug.Print "Warning: PrtDevMode is null"
        Exit Sub
    End If

    Set rpt = Nothing
    DoCmd.Close acReport, obj_name, acSaveYes
End Sub