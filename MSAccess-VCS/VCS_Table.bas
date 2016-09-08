Attribute VB_Name = "VCS_Table"
Option Compare Database

Option Private Module
Option Explicit

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