Attribute VB_Name = "VCS_Module"
Option Compare Database

Option Private Module
Option Explicit

Public Function VCS_ImportModulesFromPath(ByVal path As String) As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim fileName As String
    Dim ucs2 As Boolean

    obj_path = VCS_ObjectPath(path, "modules")

    VCS_ImportModulesFromPath = 0

    fileName = Dir$(obj_path & "*.bas")
    Do Until Len(fileName) = 0
        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        'ucs2 = VCS_File.VCS_UsingUcs2
        ucs2 = False

        If VCS_ImportExport.IsNotVCS(obj_name) Then
            VCS_IE_Functions.VCS_ImportObject acModule, obj_name, obj_path & fileName, ucs2
            VCS_ImportModulesFromPath = VCS_ImportModulesFromPath + 1
        Else
            If VCS_ImportExport.ArchiveMyself Then
                MsgBox "Module " & obj_name & " could not be updated while running. Ensure latest version is included!", vbExclamation, "Warning"
            End If
        End If

        fileName = Dir$()
    Loop
End Function

Public Function VCS_ExportModulesToPath(Db As Object, ByVal path As String) As Integer
    Dim obj_path As String
    Dim doc As Object
    Dim ucs2 As Boolean

    obj_path = VCS_ObjectPath(path, "modules")

    VCS_Dir.VCS_DeleteFilesFromDirByExtension obj_path, "bas"

    VCS_ExportModulesToPath = 0
    For Each doc In Db.Containers("Modules").Documents
        DoEvents
        If _
            (Left$(doc.name, 1) <> "~") And _
            (VCS_ImportExport.IsNotVCS(doc.name) Or VCS_ImportExport.ArchiveMyself) _
        Then
            'ucs2 = VCS_File.VCS_UsingUcs2
            ucs2 = False
            VCS_IE_Functions.VCS_ExportObject acModule, doc.name, obj_path & doc.name & ".bas", ucs2
            VCS_ExportModulesToPath = VCS_ExportModulesToPath + 1
        End If
    Next
End Function

Public Sub VCS_SanitizeModules(ByVal path As String)
    Dim obj_path As String
    obj_path = VCS_ObjectPath(path, "modules")
    VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas"
End Sub