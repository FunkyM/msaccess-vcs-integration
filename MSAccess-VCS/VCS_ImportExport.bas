Attribute VB_Name = "VCS_ImportExport"
Option Compare Database

Option Explicit

' Comma seperated list of specific tables to process with their data.
' Set to "*" to import/export contents of all tables.
Public Const IncludeTables As String = "*"

' Enables extended debug output to the immediate window
Public Const DebugOutput As Boolean = True

' Determines if VCS_ modules are to to be exported, too
Public Const ArchiveMyself As Boolean = False

' Converts UCS-2 LE encoded sources to UTF-8 automatically
' Useful since some version control systems treat UCS-2 encoded files as binary
' Not required for git
Public Const ConvertUcs2ToUtf8 As Boolean = True

Public Sub VCS_Debug(Optional ByVal strMessage As String = "", Optional ByVal strLevel As String = "Debug", Optional ByVal withNewLine As Boolean = True)
    If withNewLine Then
        strMessage = strMessage & vbNewLine
    End If

    Select Case strLevel
        Case "Info"
            Debug.Print strMessage ;
        Case "Debug"
            If DebugOutput Then
                Debug.Print strMessage ;
            End If
        Case Else
            Debug.Print strLevel & ": " & strMessage ;
    End Select
End Sub

Private Function GetElapsedTime(ByVal startTime As Single) As String
    GetElapsedTime = Round(Timer - startTime, 2) & " seconds"
End Function

' Returns true if named module is NOT part of the VCS code
Public Function IsNotVCS(ByVal name As String) As Boolean
    If _
      name <> "VCS_Dir" And _
      name <> "VCS_File" And _
      name <> "VCS_IE_Functions" And _
      name <> "VCS_ImportExport" And _
      name <> "VCS_Reference" And _
      name <> "VCS_Relation" And _
      name <> "VCS_Report" And _
      name <> "VCS_String" And _
      name <> "VCS_Table" And _
      name <> "VCS_Loader" _
    Then
        IsNotVCS = True
    Else
        IsNotVCS = False
    End If
End Function

' Exports all forms, reports, queries, macros, modules and tables to a path
Public Sub VCS_ExportAllSources(Optional ByVal sourcePath As String = vbNullString, Optional ByVal strTablesToExportWithData As String = "*")
    Dim exportPath As String
    Dim objType As Variant
    Dim objCount As Integer
    Dim startTime As Single

    startTime = Timer

    CloseFormsAndReports

    If sourcePath <> vbNullString Then
        exportPath = sourcePath
    Else
        exportPath = VCS_Dir.VCS_SourcePath()
    End If

    VCS_Dir.VCS_MkDirIfNotExist exportPath

    VCS_Debug "> Export Starting."

    For Each objType In Split("Query Form Report Macro Module Reference Table Relation")
        VCS_Debug VCS_PadRight("Exporting " & objType & " Objects...", 32), withNewLine:=False
        objCount = VCS_ExportObjects(objType, exportPath, strTablesToExportWithData)
        VCS_Debug "Done (" & objCount & ")"

        ' Sanitize queries, forms, reports and macros
        Select Case objType
            Case "Query", "Form", "Report", "Macro"
                VCS_Debug VCS_PadRight("Sanitizing " & objType & " Files...", 32), withNewLine:=False
                VCS_SanitizeFilesForObjectTypeAtPath objType, VCS_ObjectPath(exportPath, objType)
                VCS_Debug "Done"
        End Select
    Next

    VCS_Debug "> Export Finished. (" & GetElapsedTime(startTime) & ")"
End Sub

' Imports all forms, reports, queries, macros, modules and tables from a path
Public Sub VCS_ImportAllSources(Optional ByVal sourcePath As String = vbNullString)
    Dim importPath As String
    Dim FSO As Object
    Dim objType As Variant
    Dim objCount As Integer
    Dim startTime As Single

    startTime = Timer

    CloseFormsAndReports

    If sourcePath <> vbNullString Then
        importPath = sourcePath
    Else
        importPath = VCS_Dir.VCS_SourcePath()
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(importPath) Then
        MsgBox "No source found at:" & vbCrLf & importPath, vbExclamation, "Import failed"
        Exit Sub
    End If

    VCS_Debug "> Import Starting."

    For Each objType In Split("Reference Query Table Form Report Macro Module Relation")
        VCS_Debug VCS_PadRight("Importing " & objType & " Objects...", 32), withNewLine:=False
        objCount = VCS_ImportObjects(objType, importPath)
        VCS_Debug "Done (" & objCount & ")"
    Next

    DoEvents

    VCS_Debug "> Import Finished. (" & GetElapsedTime(startTime) & ")"
End Sub

' Imports all sources from a path and drops all objects
Public Sub VCS_ImportProject(Optional ByVal sourcePath As String = vbNullString, Optional ByVal prompt As Boolean = True)
    On Error GoTo errorHandler

    If _
        prompt And _
        MsgBox("This action will delete all existing: " & vbCrLf & _
              vbCrLf & _
              Chr$(149) & " Tables" & vbCrLf & _
              Chr$(149) & " Forms" & vbCrLf & _
              Chr$(149) & " Macros" & vbCrLf & _
              Chr$(149) & " Modules" & vbCrLf & _
              Chr$(149) & " Queries" & vbCrLf & _
              Chr$(149) & " Reports" & vbCrLf & _
              vbCrLf & _
              "Are you sure you want to proceed?", vbCritical + vbYesNo, _
              "Import Project") <> vbYes _
    Then
        Exit Sub
    End If

    CloseFormsAndReports

    Debug.Print "> Deleting Existing Objects"
    VCS_DeleteAllObjects

    Debug.Print "> Importing Project"
    VCS_ImportAllSources sourcePath

    Exit Sub

errorHandler:
    Debug.Print "VCS_ImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
                Err.Description
End Sub

' Close all open forms
Private Sub CloseFormsAndReports()
    On Error GoTo errorHandler
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).name
        DoEvents
    Loop
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).name
        DoEvents
    Loop
    Exit Sub

errorHandler:
    Debug.Print "VCS_ImportExport.CloseFormsAndReports: Error #" & Err.Number & vbCrLf & _
                Err.Description
End Sub