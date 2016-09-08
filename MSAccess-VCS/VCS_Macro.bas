Attribute VB_Name = "VCS_Macro"
Option Compare Database

Option Private Module
Option Explicit

Public Function VCS_ImportMacrosFromPath(ByVal path As String) As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim fileName As String
    Dim ucs2 As Boolean

    obj_path = VCS_ObjectPath(path, "macros")

    VCS_ImportMacrosFromPath = 0

    fileName = Dir$(obj_path & "*.bas")
    Do Until Len(fileName) = 0
        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        ucs2 = VCS_File.VCS_UsingUcs2
        VCS_IE_Functions.VCS_ImportObject acMacro, obj_name, obj_path & fileName, ucs2
        VCS_ImportMacrosFromPath = VCS_ImportMacrosFromPath + 1

        fileName = Dir$()
    Loop
End Function

Public Function VCS_ExportMacrosToPath(Db As Object, ByVal path As String) As Integer
    Dim obj_path As String
    Dim doc As Object
    Dim ucs2 As Boolean

    obj_path = VCS_ObjectPath(path, "macros")

    VCS_Dir.VCS_DeleteFilesFromDirByExtension obj_path, "bas"

    VCS_ExportMacrosToPath = 0
    For Each doc In Db.Containers("Scripts").Documents
        DoEvents
        If _
            (Left$(doc.name, 1) <> "~") _
        Then
            ucs2 = VCS_File.VCS_UsingUcs2
            VCS_IE_Functions.VCS_ExportObject acMacro, doc.name, obj_path & doc.name & ".bas", ucs2
            VCS_ExportMacrosToPath = VCS_ExportMacrosToPath + 1
        End If
    Next
End Function

Public Sub VCS_SanitizeMacros(ByVal path As String)
    Dim obj_path As String
    obj_path = VCS_ObjectPath(path, "macros")
    VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas"
End Sub