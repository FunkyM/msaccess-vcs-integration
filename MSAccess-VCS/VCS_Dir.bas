Attribute VB_Name = "VCS_Dir"
Option Compare Database

Option Private Module
Option Explicit

' Path/Directory of the current database file.
Public Function VCS_ProjectPath() As String
    VCS_ProjectPath = CurrentProject.Path
    If Right$(VCS_ProjectPath, 1) <> "\" Then VCS_ProjectPath = VCS_ProjectPath & "\"
End Function

Public Function VCS_SourcePath() As String
    VCS_SourcePath = VCS_ProjectPath() & CurrentProject.Name & ".src\"
End Function

Public Function VCS_AppendDirectoryDelimiter(ByVal Path As String) As String
    If Mid(Path, Len(Path), 1) = "\" Then
        VCS_AppendDirectoryDelimiter = Path
        Exit Function
    Else
        VCS_AppendDirectoryDelimiter = Path & "\"
        Exit Function
    End If
End Function

' Create folder `Path`. Silently do nothing if it already exists.
Public Sub VCS_MkDirIfNotExist(ByVal Path As String)
    On Error GoTo MkDirIfNotexist_noop
    MkDir Path
MkDirIfNotexist_noop:
    On Error GoTo 0
End Sub

' Delete a file if it exists.
Public Sub VCS_DelIfExist(ByVal Path As String)
    On Error GoTo DelIfNotExist_Noop
    Kill Path
DelIfNotExist_Noop:
    On Error GoTo 0
End Sub

' Delete all *.`ext` files in `Path`.
Public Sub VCS_DeleteFilesFromDirByExtension(ByVal Path As String, ByVal Ext As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(Path) Then Exit Sub

    On Error GoTo VCS_DeleteFilesFromDirByExtension_noop
    If Dir$(Path & "*." & Ext) <> vbNullString Then
        FSO.DeleteFile Path & "*." & Ext
    End If

VCS_DeleteFilesFromDirByExtension_noop:
    On Error GoTo 0
End Sub

Public Function VCS_FileExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    VCS_FileExists = False
    VCS_FileExists = ((GetAttr(strPath) And vbDirectory) <> vbDirectory)
End Function