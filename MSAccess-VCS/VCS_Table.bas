Attribute VB_Name = "VCS_Table"
Option Compare Database

Option Private Module
Option Explicit

Public Function VCS_ImportTableDefinitionsFromPath(ByVal path As String) As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim fileName As String

    obj_path = VCS_ObjectPath(path, "tbldef")

    VCS_ImportTableDefinitionsFromPath = 0

    fileName = Dir$(obj_path & "*.xml")
    Do Until Len(fileName) = 0
        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        VCS_ImportTableDef CStr(obj_name), obj_path
        VCS_DataMacro.VCS_ImportDataMacros obj_name, obj_path
        VCS_ImportTableDefinitionsFromPath = VCS_ImportTableDefinitionsFromPath + 1

        fileName = Dir$()
    Loop

    ' we must have access to the remote store to import these!
    fileName = Dir$(obj_path & "*.LNKD")
    Do Until Len(fileName) = 0
        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        VCS_ImportLinkedTable CStr(obj_name), obj_path & obj_name & ".LNKD"
        VCS_ImportTableDefinitionsFromPath = VCS_ImportTableDefinitionsFromPath + 1

        fileName = Dir$()
    Loop
End Function

Public Function VCS_ExportTableDefinitionsToPath(ByVal Db As Object, ByVal path As String) As Integer
    Dim obj_path As String

    obj_path = VCS_ObjectPath(path, "tbldef")

    VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
    VCS_Dir.VCS_MkDirIfNotExist obj_path

    VCS_Dir.VCS_DeleteFilesFromDirByExtension obj_path, "xml"
    VCS_Dir.VCS_DeleteFilesFromDirByExtension obj_path, "LNKD"
    VCS_Dir.VCS_DeleteFilesFromDirByExtension obj_path, "dm"

    Dim td As DAO.TableDef

    VCS_ExportTableDefinitionsToPath = 0
    For Each td In Db.TableDefs
        ' Skip system tables
        ' Skip temporary tables
        If _
            Left$(td.name, 4) <> "MSys" And _
            Left$(td.name, 1) <> "~" _
        Then
            If Len(td.connect) = 0 Then ' this is not an external table
                VCS_ExportTableDef td.name, obj_path
                VCS_DataMacro.VCS_ExportDataMacros td.name, obj_path
            Else
                VCS_ExportLinkedTable td.name, obj_path & td.name & ".LNKD"
            End If
            VCS_ExportTableDefinitionsToPath = VCS_ExportTableDefinitionsToPath + 1
        End If
    Next
End Function

Public Function VCS_ImportTablesFromPath(ByVal path As String) As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim fileName As String

    obj_path = VCS_ObjectPath(path, "tables")

    VCS_ImportTablesFromPath = 0

    fileName = Dir$(obj_path & "*.txt")
    Do Until Len(fileName) = 0
        DoEvents

        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        VCS_ImportTableData CStr(obj_name), obj_path
        VCS_ImportTablesFromPath = VCS_ImportTablesFromPath + 1

        fileName = Dir$()
    Loop
End Function

Public Function VCS_ExportTablesToPath(ByVal Db As Object, ByVal path As String, Optional ByVal IncludeTables As String = "") As Integer
    Dim obj_path As String

    obj_path = VCS_ObjectPath(path, "tables")

    VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
    VCS_Dir.VCS_DeleteFilesFromDirByExtension obj_path, "txt"

    Dim td As DAO.TableDef

    VCS_ExportTablesToPath = 0
    For Each td In Db.TableDefs
        ' Skip system tables
        ' Skip temporary tables
        If _
            Left$(td.name, 4) <> "MSys" And _
            Left$(td.name, 1) <> "~" _
        Then
            If Len(td.connect) = 0 Then ' this is not an external table
                If VCS_HasToken(IncludeTables, "*") Then
                    DoEvents
                    VCS_ExportTableData CStr(td.name), obj_path
                    If Len(Dir$(obj_path & td.name & ".txt")) > 0 Then
                        VCS_ExportTablesToPath = VCS_ExportTablesToPath + 1
                    End If
                ElseIf VCS_HasToken(IncludeTables, td.name) Then
                    DoEvents
                    On Error GoTo Err_TableNotFound
                    VCS_ExportTableData CStr(td.name), obj_path
                    VCS_ExportTablesToPath = VCS_ExportTablesToPath + 1
Err_TableNotFound:
                End If
            End If
        End If
    Next
End Function

Public Sub VCS_ExportLinkedTable(ByVal tbl_name As String, ByVal fileName As String)
    On Error GoTo Err_LinkedTable

    Dim workFileName As String
    Dim FSO As Object
    Dim OutFile As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")

    If VCS_ShouldHandleUcs2Conversion("LinkedTable") Then
        workFileName = VCS_File.VCS_TempFile()
    Else
        workFileName = fileName
    End If

    ' Open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    Set OutFile = FSO.CreateTextFile(workFileName, overwrite:=True, Unicode:=True)

    OutFile.Write CurrentDb.TableDefs(tbl_name).name
    OutFile.Write vbCrLf

    If InStr(1, CurrentDb.TableDefs(tbl_name).connect, "DATABASE=" & CurrentProject.Path) Then
        ' Change to relative path
        Dim connect() As String
        connect = Split(CurrentDb.TableDefs(tbl_name).connect, CurrentProject.Path)
        OutFile.Write connect(0) & "." & connect(1)
    Else
        OutFile.Write CurrentDb.TableDefs(tbl_name).connect
    End If

    OutFile.Write vbCrLf
    OutFile.Write CurrentDb.TableDefs(tbl_name).SourceTableName
    OutFile.Write vbCrLf

    Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim td As DAO.TableDef
    Set td = Db.TableDefs(tbl_name)
    Dim idx As DAO.Index

    For Each idx In td.Indexes
        If idx.Primary Then
            OutFile.Write Right$(idx.Fields, Len(idx.Fields) - 1)
            OutFile.Write vbCrLf
        End If

    Next

Err_LinkedTable_Fin:
    On Error Resume Next
    OutFile.Close

    If VCS_ShouldHandleUcs2Conversion("LinkedTable") Then
        ' Save files as .odbc
        VCS_File.VCS_ConvertUcs2Utf8 workFileName, fileName
    End If

    Exit Sub

Err_LinkedTable:
    OutFile.Close
    MsgBox Err.Description, vbCritical, "ERROR: EXPORT LINKED TABLE"
    Resume Err_LinkedTable_Fin
End Sub

' Save a Table Definition as SQL statement
Public Sub VCS_ExportTableDef(ByVal TableName As String, ByVal directory As String)
    Dim fileName As String
    fileName = directory & TableName & ".xml"
    
    Application.ExportXML _
    ObjectType:=acExportTable, _
    DataSource:=TableName, _
    SchemaTarget:=fileName
    
    'exort Data Macros
    VCS_DataMacro.VCS_ExportDataMacros TableName, directory
End Sub


' Determine if a table or exists.
' based on sample code of support.microsoftcom
' ARGUMENTS:
'    TName: The name of a table or query.
'
' RETURNS: True (it exists) or False (it does not exist).
Private Function TableExists(ByVal TName As String) As Boolean
    Dim Db As DAO.Database
    Dim Found As Boolean
    Dim Test As String

    Const NAME_NOT_IN_COLLECTION As Integer = 3265

     ' Assume the table or query does not exist.
    Found = False
    Set Db = CurrentDb()

     ' Trap for any errors.
    On Error Resume Next

     ' See if the name is in the Tables collection.
    Test = Db.TableDefs(TName).name
    If Err.Number <> NAME_NOT_IN_COLLECTION Then Found = True

    ' Reset the error variable.
    Err = 0

    TableExists = Found
End Function

' Export the lookup table `tblName` to `source\tables`.
Public Sub VCS_ExportTableData(ByVal tbl_name As String, ByVal obj_path As String)
    Dim FSO As Object
    Dim OutFile As Object
    Dim rs As DAO.Recordset
    Dim fieldObj As Object
    Dim c As Long, Value As Variant

    ' Check first
    If Not TableExists(tbl_name) Then
        Debug.Print "Error: Table " & tbl_name & " missing"
        Exit Sub
    End If

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & tbl_name)
    If rs.RecordCount = 0 Then
        rs.Close
        Exit Sub
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    VCS_Dir.VCS_MkDirIfNotExist obj_path

    Dim tempFileName As String
    tempFileName = VCS_File.VCS_TempFile()

    ' Open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    Set OutFile = FSO.CreateTextFile(tempFileName, overwrite:=True, Unicode:=True)

    c = 0
    For Each fieldObj In rs.Fields
        If c <> 0 Then OutFile.Write vbTab
        c = c + 1
        OutFile.Write fieldObj.name
    Next
    OutFile.Write vbCrLf

    rs.MoveFirst
    Do Until rs.EOF
        c = 0
        For Each fieldObj In rs.Fields
            DoEvents
            If c <> 0 Then OutFile.Write vbTab
            c = c + 1
            Value = rs(fieldObj.name)
            If IsNull(Value) Then
                Value = vbNullString
            Else
                Value = Replace(Value, "\", "\\")
                Value = Replace(Value, vbCrLf, "\n")
                Value = Replace(Value, vbCr, "\n")
                Value = Replace(Value, vbLf, "\n")
                Value = Replace(Value, vbTab, "\t")
            End If
            OutFile.Write Value
        Next
        OutFile.Write vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    OutFile.Close

    VCS_File.VCS_ConvertUcs2Utf8 tempFileName, obj_path & tbl_name & ".txt"
    FSO.DeleteFile tempFileName
End Sub

Public Sub VCS_ImportLinkedTable(ByVal tblName As String, ByVal fileName As String)
    Dim workFileName As String
    Dim Db As DAO.Database
    Dim FSO As Object
    Dim InFile As Object

    Set Db = CurrentDb()
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If VCS_ShouldHandleUcs2Conversion("LinkedTable") Then
        workFileName = VCS_File.VCS_TempFile()
        VCS_ConvertUtf8Ucs2 fileName, workFileName
    Else
        workFileName = fileName
    End If

    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(workFileName, iomode:=ForReading, create:=False, Format:=TristateTrue)

    On Error GoTo err_notable:
    DoCmd.DeleteObject acTable, tblName

    GoTo err_notable_fin

err_notable:
    Err.Clear
    Resume err_notable_fin

err_notable_fin:
    On Error GoTo Err_CreateLinkedTable:

    Dim td As DAO.TableDef
    Set td = Db.CreateTableDef(InFile.ReadLine())

    Dim connect As String
    connect = InFile.ReadLine()
    If InStr(1, connect, "DATABASE=.\") Then 'replace relative path with literal path
        connect = Replace(connect, "DATABASE=.\", "DATABASE=" & VCS_AppendDirectoryDelimiter(CurrentProject.Path))
    End If
    td.connect = connect

    td.SourceTableName = InFile.ReadLine()
    Db.TableDefs.Append td

    GoTo Err_CreateLinkedTable_Fin

Err_CreateLinkedTable:
    MsgBox Err.Description, vbCritical, "ERROR: IMPORT LINKED TABLE"
    Resume Err_CreateLinkedTable_Fin

Err_CreateLinkedTable_Fin:
    ' This will throw errors if a primary key already exists or the table is linked to an access database table
    ' This will also error out if no pk is present
    On Error GoTo Err_LinkPK_Fin:

    Dim Fields As String
    Fields = InFile.ReadLine()
    Dim Field As Variant
    Dim sql As String
    sql = "CREATE INDEX __uniqueindex ON " & td.name & " ("

    For Each Field In Split(Fields, ";+")
        sql = sql & "[" & Field & "]" & ","
    Next
    ' Remove extraneous comma
    sql = Left$(sql, Len(sql) - 1)

    sql = sql & ") WITH PRIMARY"

    Db.Execute sql

Err_LinkPK_Fin:
    On Error Resume Next
    InFile.Close
End Sub

' Import Table Definition
Public Sub VCS_ImportTableDef(ByVal tblName As String, ByVal directory As String)
    Dim filePath As String
    
    filePath = directory & tblName & ".xml"
    Application.ImportXML DataSource:=filePath, ImportOptions:=acStructureOnly

End Sub

' Import the lookup table `tblName` from `source\tables`.
Public Sub VCS_ImportTableData(ByVal tblName As String, ByVal obj_path As String)
    Dim Db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim FSO As Object
    Dim InFile As Object
    Dim c As Long, buf As String, Values() As String, Value As Variant

    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim tempFileName As String
    tempFileName = VCS_File.VCS_TempFile()
    VCS_File.VCS_ConvertUtf8Ucs2 obj_path & tblName & ".txt", tempFileName

    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, iomode:=ForReading, create:=False, Format:=TristateTrue)
    Set Db = CurrentDb

    Db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = Db.OpenRecordset(tblName)
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = InFile.ReadLine()
        If Len(Trim$(buf)) > 0 Then
            Values = Split(buf, vbTab)
            c = 0
            rs.AddNew
            For Each fieldObj In rs.Fields
                DoEvents
                Value = Values(c)
                If Len(Value) = 0 Then
                    Value = Null
                Else
                    Value = Replace(Value, "\t", vbTab)
                    Value = Replace(Value, "\n", vbCrLf)
                    Value = Replace(Value, "\\", "\")
                End If
                rs(fieldObj.name) = Value
                c = c + 1
            Next
            rs.Update
        End If
    Loop

    rs.Close
    InFile.Close
    FSO.DeleteFile tempFileName
End Sub