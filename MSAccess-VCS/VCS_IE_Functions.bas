Attribute VB_Name = "VCS_IE_Functions"
Option Compare Database

Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Private Const AggressiveSanitize As Boolean = True
Private Const StripPublishOption As Boolean = True

' Constants for Scripting.FileSystemObject API
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8
Public Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2

' Returns the file extension for an object type
' See https://msdn.microsoft.com/en-us/library/aa241721(v=vs.60).aspx
Public Function VCS_GetObjectFileExtension(ByVal objType As String) As String
    Select Case objType
        Case "Reference"
            VCS_GetObjectFileExtension = "csv"
        Case "Query"
            VCS_GetObjectFileExtension = "bas"
        Case "Table"
            VCS_GetObjectFileExtension = "xml"
        Case "TableDefinition"
            VCS_GetObjectFileExtension = "xsd"
        Case "LinkedTable"
            VCS_GetObjectFileExtension = "lnkd"
        Case "TableDataMacro"
            VCS_GetObjectFileExtension = "xml"
        Case "Form"
            VCS_GetObjectFileExtension = "bas"
        Case "Report"
            VCS_GetObjectFileExtension = "bas"
        Case "Macro"
            VCS_GetObjectFileExtension = "bas"
        Case "Module"
            VCS_GetObjectFileExtension = "bas"
        Case "Relation"
            VCS_GetObjectFileExtension = "txt"
        Case "PrintVars"
            VCS_GetObjectFileExtension = "txt"
        Case "DatabaseProperties"
            VCS_GetObjectFileExtension = "txt"
    End Select
End Function

' Returns a path to store files for a specific object type
Public Function VCS_ObjectPath(ByVal strPath As String, ByVal strObjectType As String) As String
    Select Case strObjectType
        Case "TableDefinition", "LinkedTable", "TableDataMacro"
            VCS_ObjectPath = "tbldef"
        Case "Table"
            VCS_ObjectPath = "tables"
        Case "Report", "PrintVars"
            VCS_ObjectPath = "reports"
        Case "Macro"
            VCS_ObjectPath = "macros"
        Case "Module"
            VCS_ObjectPath = "modules"
        Case "Query"
            VCS_ObjectPath = "queries"
        Case "Relation"
            VCS_ObjectPath = "relations"
        Case "Form"
            VCS_ObjectPath = "forms"
        Case "References"
            VCS_ObjectPath = "\"
        Case "DatabaseProperties"
            VCS_ObjectPath = "\"
        Case Else
            VCS_ObjectPath = ""
    End Select
    VCS_ObjectPath = strPath & VCS_ObjectPath
End Function

Public Function VCS_ShouldHandleUcs2Conversion(ByVal objType As String) As Boolean
    If _
        VCS_UsingUcs2() And _
        VCS_ImportExport.ConvertUcs2ToUtf8 _
    Then
        Select Case objType
            Case "Query", "Form", "Report", "Macro", "TableDataMacro"
                VCS_ShouldHandleUcs2Conversion = True
            Case "Module"
                ' Modules always use UTF-8
                VCS_ShouldHandleUcs2Conversion = False
            Case Else
                VCS_ShouldHandleUcs2Conversion = False
        End Select
    Else
        VCS_ShouldHandleUcs2Conversion = False
    End If
End Function

' Export a database object with optional UCS2-to-UTF-8 conversion.
Public Sub VCS_ExportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                    ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False)

    VCS_Dir.VCS_MkDirIfNotExist Left$(file_path, InStrRev(file_path, "\"))

    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        Application.SaveAsText obj_type_num, obj_name, tempFileName
        VCS_File.VCS_ConvertUcs2Utf8 tempFileName, file_path
    Else
        Application.SaveAsText obj_type_num, obj_name, file_path
    End If
End Sub

' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub VCS_ImportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                    ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False)

    If Not VCS_Dir.VCS_FileExists(file_path) Then Exit Sub

    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        VCS_File.VCS_ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName

        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        FSO.DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

Public Sub VCS_DeleteAllObjects()
    Dim objType As Variant

    For Each objType In Split("Relation Query Table Form Report Macro Module")
        VCS_DeleteObjects(objType)
    Next
End Sub

' Deletes all objects of given type from the current database
Public Sub VCS_DeleteObjects(ByVal objType As String)
    Dim dbSource As DAO.Database
    Dim doc As Object

    Set dbSource = CurrentDb()

    Select Case objType
        Case "Relation"
            ' Delete Relations
            For Each doc In dbSource.Relations
                If _
                    Not (doc.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or _
                    doc.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") _
                Then
                    dbSource.Relations.Delete(doc.name)
                End If
            Next
        Case "Query"
            ' Delete Queries
            For Each doc In dbSource.QueryDefs
                DoEvents
                If _
                    Left$(doc.name, 1) <> "~" _
                Then
                    dbSource.QueryDefs.Delete(doc.name)
                End If
            Next
        Case "Table"
            ' Delete Table Definitions
            For Each doc In dbSource.TableDefs
                If _
                    Left$(doc.name, 4) <> "MSys" And _
                    Left$(doc.name, 1) <> "~" _
                Then
                    dbSource.TableDefs.Delete(doc.name)
                End If
            Next
        Case "Form", "Report", "Macro", "Module"
            ' Delete Forms, Reports, Macros and Modules
            DoEvents
            For Each doc In dbSource.Containers(VCS_GetContainerNameForObjectType(objType)).Documents
                DoEvents
                If _
                    (Left$(doc.name, 1) <> "~") _
                Then
                    If _
                        objType <> "Module" Or _
                        (objType = "Module" And IsNotVCS(doc.name)) _
                    Then
                        DoCmd.DeleteObject VCS_GetAccessTypeForObjectType(objType), doc.name
                    End If
                End If
            Next
        Case ""
    End Select
End Sub

' For each *.`ext` in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Public Sub VCS_SanitizeTextFiles(ByVal Path As String, ByVal Ext As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Setup Block matching Regex.
    Dim rxBlock As Object
    Set rxBlock = CreateObject("VBScript.RegExp")
    rxBlock.ignoreCase = False

    ' Match PrtDevNames / Mode with or without W
    Dim srchPattern As String
    srchPattern = "PrtDev(?:Names|Mode)[W]?"
    If (AggressiveSanitize = True) Then
        '  Add and group aggressive matches
        srchPattern = "(?:" & srchPattern
        srchPattern = srchPattern & "|GUID|""GUID""|NameMap|dbLongBinary ""DOL"""
        srchPattern = srchPattern & ")"
    End If

    ' Ensure that this is the begining of a block.
    srchPattern = srchPattern & " = Begin"
    rxBlock.Pattern = srchPattern

    ' Setup Line Matching Regex.
    Dim rxLine As Object
    Set rxLine = CreateObject("VBScript.RegExp")
    srchPattern = "^\s*(?:"
    srchPattern = srchPattern & "Checksum ="
    srchPattern = srchPattern & "|BaseInfo|NoSaveCTIWhenDisabled =1"
    If (StripPublishOption = True) Then
        srchPattern = srchPattern & "|dbByte ""PublishToWeb"" =""1"""
        srchPattern = srchPattern & "|PublishOption =1"
    End If
    srchPattern = srchPattern & ")"
    rxLine.Pattern = srchPattern

    Dim fileName As String
    fileName = Dir$(Path & "*." & Ext)

    Dim isReport As Boolean
    isReport = False

    Do Until Len(fileName) = 0
        DoEvents
        Dim obj_name As String
        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        Dim InFile As Object
        Set InFile = FSO.OpenTextFile(Path & obj_name & "." & Ext, iomode:=ForReading, create:=False, Format:=TristateFalse)
        Dim OutFile As Object
        Set OutFile = FSO.CreateTextFile(Path & obj_name & ".sanitize", overwrite:=True, Unicode:=False)

        Dim getLine As Boolean
        getLine = True

        Do Until InFile.AtEndOfStream
            DoEvents
            Dim txt As String

            ' Check if we need to get a new line of text
            If getLine = True Then
                txt = InFile.ReadLine
            Else
                getLine = True
            End If

            ' Skip lines starting with line pattern
            If rxLine.Test(txt) Then
                Dim rxIndent As Object
                Set rxIndent = CreateObject("VBScript.RegExp")
                rxIndent.Pattern = "^(\s+)\S"

                ' Get indentation level.
                Dim matches As Object
                Set matches = rxIndent.Execute(txt)

                ' Setup pattern to match current indent
                Select Case matches.Count
                    Case 0
                        rxIndent.Pattern = "^" & vbNullString
                    Case Else
                        rxIndent.Pattern = "^" & matches(0).SubMatches(0)
                End Select
                rxIndent.Pattern = rxIndent.Pattern + "\s"

                ' Skip lines with deeper indentation
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If Not rxIndent.Test(txt) Then Exit Do
                Loop
                ' We've moved on at least one line so do get a new one
                ' when starting the loop again.
                getLine = False

            ' skip blocks of code matching block pattern
            ElseIf rxBlock.Test(txt) Then
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            ElseIf InStr(1, txt, "Begin Report") = 1 Then
                isReport = True
                OutFile.WriteLine txt
            ElseIf isReport = True And (InStr(1, txt, "    Right =") Or InStr(1, txt, "    Bottom =")) Then
                'skip line
                If InStr(1, txt, "    Bottom =") Then
                    isReport = False
                End If
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        InFile.Close

        FSO.DeleteFile (Path & fileName)

        Dim thisFile As Object
        Set thisFile = FSO.GetFile(Path & obj_name & ".sanitize")

        ' Error Handling to deal with errors caused by Dropbox, VirusScan,
        ' or anything else touching the file.
        Dim ErrCounter As Integer
        On Error GoTo ErrorHandler
        thisFile.Move (Path & fileName)
        fileName = Dir$()
    Loop

    Exit Sub
ErrorHandler:
    ErrCounter = ErrCounter + 1
    If ErrCounter = 20 Then  ' 20 attempts seems like a nice arbitrary number
        MsgBox "This file could not be moved: " & vbNewLine, vbCritical + vbApplicationModal, _
            "Error moving file..."
        Resume Next
    End If
    Select Case Err.Number
        Case 58    ' "File already exists" error.
            DoEvents
            Sleep 10
            Resume    ' Go back to what you were doing
        Case Else
            MsgBox "This file could not be moved: " & vbNewLine, vbCritical + vbApplicationModal, _
                "Error moving file..."
    End Select
    Resume Next
End Sub