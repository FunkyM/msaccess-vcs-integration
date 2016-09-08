Attribute VB_Name = "VCS_Form"
Option Compare Database

Option Private Module
Option Explicit

Public Function VCS_ImportFormsFromPath(ByVal path As String) As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim fileName As String
    Dim ucs2 As Boolean

    obj_path = VCS_ObjectPath(path, "forms")

    VCS_ImportFormsFromPath = 0

    fileName = Dir$(obj_path & "*.bas")
    Do Until Len(fileName) = 0
        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        ucs2 = VCS_File.VCS_UsingUcs2
        VCS_IE_Functions.VCS_ImportObject acForm, obj_name, obj_path & fileName, ucs2
        VCS_ImportFormsFromPath = VCS_ImportFormsFromPath + 1

        fileName = Dir$()
    Loop
End Function

Public Function VCS_ExportFormsToPath(Db As Object, ByVal path As String) As Integer
    Dim obj_path As String
    Dim doc As Object ' DAO.Document
    Dim ucs2 As Boolean

    obj_path = VCS_ObjectPath(path, "forms")

    VCS_Dir.VCS_DeleteFilesFromDirByExtension obj_path, "bas"

    VCS_ExportFormsToPath = 0
    For Each doc In Db.Containers("Forms").Documents
        DoEvents
        If _
            (Left$(doc.name, 1) <> "~") _
        Then
            ucs2 = VCS_File.VCS_UsingUcs2
            VCS_IE_Functions.VCS_ExportObject acForm, doc.name, obj_path & doc.name & ".bas", ucs2
            VCS_ExportFormsToPath = VCS_ExportFormsToPath + 1
        End If
    Next
End Function

Public Sub VCS_SanitizeForms(ByVal path As String)
    Dim obj_path As String
    obj_path = VCS_ObjectPath(path, "forms")
    VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas"
End Sub
