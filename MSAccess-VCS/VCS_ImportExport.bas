Attribute VB_Name = "VCS_ImportExport"
Option Compare Database

Option Explicit

' Comma seperated list of specific tables to process with their data.
' Set to "*" to import/export contents of all tables.
Public Const IncludeTables As String = "*"

' Enables extended debug output to the immediate window
Public Const DebugOutput As Boolean = False

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
      name <> "VCS_DataMacro" And _
      name <> "VCS_Dir" And _
      name <> "VCS_File" And _
      name <> "VCS_Form" And _
      name <> "VCS_IE_Functions" And _
      name <> "VCS_ImportExport" And _
      name <> "VCS_Macro" And _
      name <> "VCS_Module" And _
      name <> "VCS_Query" And _
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
Public Sub VCS_ExportAllSources(Optional ByVal sourcePath As String = vbNullString, Optional IncludeTables As String = "")
    Dim Db As Object ' DAO.Database
    Dim source_path As String
    Dim procToCall As String
    Dim obj As Variant
    Dim obj_count As Integer
    Dim result As Variant
    Dim startTime As Single

    startTime = Timer

    Set Db = CurrentDb()

    CloseFormsAndReports

    If sourcePath <> vbNullString Then
        source_path = sourcePath
    Else
        source_path = VCS_Dir.VCS_SourcePath()
    End If

    VCS_Dir.VCS_MkDirIfNotExist source_path

    For Each obj In Split("Queries Forms Reports Macros Modules References TableDefinitions Tables Relations")
        Debug.Print VCS_PadRight("Exporting " & obj & "...", 24);

        procToCall = "VCS_Export" & obj & "ToPath"

        Select Case obj
            Case "Tables"
                obj_count = Application.Run(procToCall, Db, source_path, IncludeTables)
            Case "References"
                obj_count = Application.Run(procToCall, source_path)
            Case Else
                obj_count = Application.Run(procToCall, Db, source_path)
        End Select

        Debug.Print "[" & obj_count & "]"

        procToCall = "VCS_Sanitize" & obj

        ' Sanitize some objects
        Select Case obj
            Case "Queries", "Forms", "Reports", "Macros"
                Debug.Print VCS_PadRight("Sanitizing " & obj & "...", 15);
                result = Application.Run(procToCall, source_path)
                Debug.Print "Done."
        End Select
    Next

    Debug.Print "Done. (" & GetElapsedTime(startTime) & ")"
End Sub

' Imports all forms, reports, queries, macros, modules and tables from a path
Public Sub VCS_ImportAllSources(Optional ByVal sourcePath As String = vbNullString)
    Dim FSO As Object
    Dim source_path As String
    Dim procToCall As String
    Dim obj As Variant
    Dim obj_count As Integer
    Dim startTime As Single

    startTime = Timer

    CloseFormsAndReports

    If sourcePath <> vbNullString Then
        source_path = sourcePath
    Else
        source_path = VCS_Dir.VCS_SourcePath()
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    For Each obj In Split("References Queries TableDefinitions Tables Forms Reports Macros Modules Relations")
        Debug.Print VCS_PadRight("Importing " & obj & "...", 24);

        procToCall = "VCS_Import" & obj & "FromPath"

        obj_count = Application.Run(procToCall, source_path)

        Debug.Print "[" & obj_count & "]"
    Next

    DoEvents

    Debug.Print "Done. (" & GetElapsedTime(startTime) & ")"
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

    Debug.Print
    Debug.Print "Deleting Existing Objects"
    Debug.Print

    VCS_DeleteAllObjects

    Debug.Print "================="
    Debug.Print "Importing Project"

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