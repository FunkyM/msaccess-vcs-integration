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

' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
Public Sub ExportAllSource()
    Dim Db As Object ' DAO.Database
    Dim source_path As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim obj_data_count As Integer
    Dim ucs2 As Boolean

    Set Db = CurrentDb

    CloseFormsAndReports

    source_path = VCS_Dir.VCS_ProjectPath() & "source\"
    VCS_Dir.VCS_MkDirIfNotExist source_path

    Debug.Print

    obj_path = source_path & "queries\"
    VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "bas"
    Debug.Print VCS_String.VCS_PadRight("Exporting queries...", 24);
    obj_count = 0
    For Each qry In Db.QueryDefs
        DoEvents
        If Left$(qry.name, 1) <> "~" Then
            VCS_IE_Functions.VCS_ExportObject acQuery, qry.name, obj_path & qry.name & ".bas", VCS_File.VCS_UsingUcs2
            obj_count = obj_count + 1
        End If
    Next
    Debug.Print VCS_String.VCS_PadRight("Sanitizing...", 15);
    VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas"
    Debug.Print "[" & obj_count & "]"

    
    For Each obj_type In Split( _
        "forms|Forms|" & acForm & "," & _
        "reports|Reports|" & acReport & "," & _
        "macros|Scripts|" & acMacro & "," & _
        "modules|Modules|" & acModule _
        , "," _
    )
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        obj_type_name = obj_type_split(1)
        obj_type_num = Val(obj_type_split(2))
        obj_path = source_path & obj_type_label & "\"
        obj_count = 0
        VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "bas"
        Debug.Print VCS_String.VCS_PadRight("Exporting " & obj_type_label & "...", 24);
        For Each doc In Db.Containers(obj_type_name).Documents
            DoEvents
            If (Left$(doc.name, 1) <> "~") And _
               (IsNotVCS(doc.name) Or ArchiveMyself) Then
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = VCS_File.VCS_UsingUcs2
                End If
                VCS_IE_Functions.VCS_ExportObject obj_type_num, doc.name, obj_path & doc.name & ".bas", ucs2
                
                If obj_type_label = "reports" Then
                    VCS_Report.VCS_ExportPrintVars doc.name, obj_path & doc.name & ".pv"
                End If
                
                obj_count = obj_count + 1
            End If
        Next

		Debug.Print VCS_String.VCS_PadRight("Sanitizing...", 15);
        If obj_type_label <> "modules" Then
            VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas"
        End If
        Debug.Print "[" & obj_count & "]"
    Next
    
    VCS_Reference.VCS_ExportReferences source_path

'-------------------------table export------------------------
    obj_path = source_path & "tables\"
    VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
    VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "txt"
    
    Dim td As DAO.TableDef
    Dim tds As DAO.TableDefs
    Set tds = Db.TableDefs

    obj_type_label = "tbldef"
    obj_type_name = "Table_Def"
    obj_type_num = acTable
    obj_path = source_path & obj_type_label & "\"
    obj_count = 0
    obj_data_count = 0
    VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
    
    'move these into Table and DataMacro modules?
    ' - We don't want to determin file extentions here - or obj_path either!
    VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "sql"
    VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "xml"
    VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "LNKD"
    
    Dim IncludeTablesCol As Collection
    Set IncludeTablesCol = StrSetToCol(INCLUDE_TABLES, ",")
    
    Debug.Print VCS_String.VCS_PadRight("Exporting " & obj_type_label & "...", 24);
    
    For Each td In tds
        ' This is not a system table
        ' this is not a temporary table
        If Left$(td.name, 4) <> "MSys" And _
        Left$(td.name, 1) <> "~" Then
            If Len(td.connect) = 0 Then ' this is not an external table
                VCS_Table.VCS_ExportTableDef td.name, obj_path
                If INCLUDE_TABLES = "*" Then
                    DoEvents
                    VCS_Table.VCS_ExportTableData CStr(td.name), source_path & "tables\"
                    If Len(Dir$(source_path & "tables\" & td.name & ".txt")) > 0 Then
                        obj_data_count = obj_data_count + 1
                    End If
                ElseIf (Len(Replace(INCLUDE_TABLES, " ", vbNullString)) > 0) And INCLUDE_TABLES <> "*" Then
                    DoEvents
                    On Error GoTo Err_TableNotFound
                    If InCollection(IncludeTablesCol,td.name) Then
                        VCS_Table.VCS_ExportTableData CStr(td.name), source_path & "tables\"
                        obj_data_count = obj_data_count + 1
                    End If
Err_TableNotFound:
                    
                'else don't export table data
                End If
            Else
                VCS_Table.VCS_ExportLinkedTable td.name, obj_path
            End If
            
            obj_count = obj_count + 1
            
        End If
    Next
    Debug.Print "[" & obj_count & "]"
    If obj_data_count > 0 Then
      Debug.Print VCS_String.VCS_PadRight("Exported data...", 24) & "[" & obj_data_count & "]"
    End If
    
    
    Debug.Print VCS_String.VCS_PadRight("Exporting Relations...", 24);
    obj_count = 0
    obj_path = source_path & "relations\"
    VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))

    VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "txt"

    Dim aRelation As DAO.Relation
    
    For Each aRelation In CurrentDb.Relations
        ' Exclude relations from system tables and inherited (linked) relations
        ' Skip if dbRelationDontEnforce property is not set. The relationship is already in the table xml file. - sean
        If Not (aRelation.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" _
                Or aRelation.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups" _
                Or (aRelation.Attributes And DAO.RelationAttributeEnum.dbRelationInherited) = _
                DAO.RelationAttributeEnum.dbRelationInherited) _
                And (aRelation.Attributes = DAO.RelationAttributeEnum.dbRelationDontEnforce) Then
            VCS_Relation.VCS_ExportRelation aRelation, obj_path & aRelation.name & ".txt"
            obj_count = obj_count + 1
        End If
    Next
    Debug.Print "[" & obj_count & "]"
    
    Debug.Print "Done."
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