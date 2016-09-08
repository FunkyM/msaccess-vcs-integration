Attribute VB_Name = "VCS_Query"
Option Compare Database

Option Private Module
Option Explicit

Public Function VCS_ImportQueriesFromPath(ByVal path As String) As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim fileName As String
    Dim ucs2 As Boolean

    obj_path = VCS_ObjectPath(path, "queries")

    VCS_ImportQueriesFromPath = 0

    Dim tempFilePath As String
    tempFilePath = VCS_File.VCS_TempFile()

    fileName = Dir$(obj_path & "*.bas")
    Do Until Len(fileName) = 0
        DoEvents

        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        ucs2 = VCS_File.VCS_UsingUcs2
        VCS_IE_Functions.VCS_ImportObject acQuery, obj_name, obj_path & fileName, ucs2
        'VCS_IE_Functions.VCS_ExportObject acQuery, obj_name, tempFilePath, ucs2
        'VCS_IE_Functions.VCS_ImportObject acQuery, obj_name, tempFilePath, ucs2
        VCS_ImportQueriesFromPath = VCS_ImportQueriesFromPath + 1

        fileName = Dir$()
    Loop

    VCS_Dir.VCS_DelIfExist tempFilePath
End Function

Public Function VCS_ExportQueriesToPath(Db As Object, ByVal path As String) As Integer
    Dim obj_path As String
    Dim doc As Object
    Dim ucs2 As Boolean

    obj_path = VCS_ObjectPath(path, "queries")

    VCS_Dir.VCS_DeleteFilesFromDirByExtension obj_path, "bas"

    VCS_ExportQueriesToPath = 0
    For Each doc In Db.QueryDefs
        DoEvents
        If _
            (Left$(doc.name, 1) <> "~") _
        Then
            ucs2 = VCS_File.VCS_UsingUcs2
            VCS_IE_Functions.VCS_ExportObject acQuery, doc.name, obj_path & doc.name & ".bas", ucs2
            VCS_ExportQueriesToPath = VCS_ExportQueriesToPath + 1
        End If
    Next
End Function

Public Sub VCS_SanitizeQueries(ByVal path As String)
    Dim obj_path As String
    obj_path = VCS_ObjectPath(path, "queries")
    VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas"
End Sub