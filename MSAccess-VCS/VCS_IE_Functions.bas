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

Public Function VCS_GetAccessTypeForObjectType(ByVal objType As String) As Integer
    Select Case objType
        Case "Query"
            VCS_GetAccessTypeForObjectType = acQuery
        Case "Form"
            VCS_GetAccessTypeForObjectType = acForm
        Case "Report"
            VCS_GetAccessTypeForObjectType = acReport
        Case "Macro"
            VCS_GetAccessTypeForObjectType = acMacro
        Case "Module"
            VCS_GetAccessTypeForObjectType = acModule
        Case "TableDataMacro"
            VCS_GetAccessTypeForObjectType = acTableDataMacro
        Case "DatabaseProperties"
            VCS_GetAccessTypeForObjectType = acDatabaseProperties
    End Select
End Function

Public Function VCS_GetContainerNameForObjectType(ByVal objType As String) As String
    Select Case objType
        Case "Form"
            VCS_GetContainerNameForObjectType = "Forms"
        Case "Report"
            VCS_GetContainerNameForObjectType = "Reports"
        Case "Macro"
            VCS_GetContainerNameForObjectType = "Scripts"
        Case "Module"
            VCS_GetContainerNameForObjectType = "Modules"
        Case Else
            VCS_GetContainerNameForObjectType = ""
    End Select
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

Public Sub VCS_LoadAccessTypeFromText(ByVal objType As String, ByVal objName As String, ByVal fileName As String)
    If VCS_ShouldHandleUcs2Conversion(objType) Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        VCS_File.VCS_ConvertUtf8Ucs2 fileName, tempFileName
        Application.LoadFromText VCS_GetAccessTypeForObjectType(objType), objName, tempFileName
        VCS_DelIfExist tempFileName
    Else
        Application.LoadFromText VCS_GetAccessTypeForObjectType(objType), objName, fileName
    End If
End Sub

Public Sub VCS_LoadLinkedTableFromText(ByVal objName As String, ByVal fileName As String)
    VCS_ImportLinkedTable objName, fileName
End Sub

Public Sub VCS_LoadTableFromText(ByVal objName As String, ByVal fileName As String)
    Application.ImportXML _
        DataSource := fileName, _
        ImportOptions := acAppendData
End Sub

Public Sub VCS_LoadTableDefinitionFromText(ByVal objName As String, ByVal fileName As String)
    On Error Resume Next
    Dim dbSource As Object
    Set dbSource = CurrentDb()
    dbSource.Execute "DROP TABLE [" & objName & "]"
    On Error Goto 0

    Application.ImportXML _
        DataSource := fileName, _
        ImportOptions := acStructureOnly
End Sub

Public Sub VCS_LoadReferencesFromText(ByVal fileName As String)
    VCS_ImportReferencesFromPath(fileName)
End Sub

Public Sub VCS_LoadRelationFromText(ByVal fileName As String)
    VCS_ImportRelation(fileName)
End Sub

Public Sub VCS_LoadFromText(ByVal objType As String, ByVal objName As String, ByVal fileName As String)
    If Not VCS_FileExists(fileName) Then Exit Sub

    Select Case objType
        Case "Query", "Form", "Report", "Macro", "Module", "TableDataMacro"
            VCS_LoadAccessTypeFromText objType, objName, fileName
        Case "LinkedTable"
            VCS_LoadLinkedTableFromText objName, fileName
        Case "Table"
            VCS_LoadTableFromText objName, fileName
        Case "TableDefinition"
            VCS_LoadTableDefinitionFromText objName, fileName
        Case "Reference"
            VCS_LoadReferencesFromText fileName
        Case "Relation"
            VCS_LoadRelationFromText fileName
        Case "PrintVars"
            VCS_ImportPrintVars objName, fileName
    End Select
End Sub

Public Sub VCS_ImportObject(ByVal objType As String, ByVal objName As String, ByVal strPath As String)
    Dim strObjectFileName As String
    strObjectFileName = VCS_ObjectNameToPathName(objName) & "." & VCS_GetObjectFileExtension(objType)
    VCS_LoadFromText objType, objName, VCS_AppendDirectoryDelimiter(strPath) & strObjectFileName
End Sub

Public Function VCS_ImportObjects(ByVal objType As String, ByVal strPath As String)
    Dim objName As String
    Dim objPath As String
    Dim fileName As String

    VCS_ImportObjects = 0
    objPath = VCS_ObjectPath(strPath, objType)

    ' Deserialize
    Select Case objType
        Case "Reference"
            VCS_ImportObject objType, "references", objPath
            VCS_ImportObjects = VCS_ImportObjects + 1
        Case "Query", "Table", "Form", "Report", "Macro", "Relation", "Module"
            fileName = Dir$(VCS_AppendDirectoryDelimiter(objPath) & "*." & VCS_GetObjectFileExtension(objType))
            Do Until Len(fileName) = 0
                objName = VCS_PathNameToObjectName(Mid$(fileName, 1, InStrRev(fileName, ".") - 1))
                Select Case objType
                    Case "Table"
                        ' Skip system tables and temporary tables
                        If _
                            Left$(objName, 4) <> "MSys" And _
                            Left$(objName, 1) <> "~" _
                        Then
                            VCS_ImportObject "TableDefinition", objName, VCS_ObjectPath(strPath, "TableDefinition")
                            VCS_ImportObject "TableDataMacro", objName, VCS_ObjectPath(strPath, "TableDataMacro")
                            VCS_ImportObject objType, objName, objPath
                            VCS_ImportObjects = VCS_ImportObjects + 1
                        End If
                    Case "Report"
                        VCS_ImportObject objType, objName, objPath
                        VCS_ImportObject "PrintVars", objName, VCS_ObjectPath(strPath, "PrintVars")
                        VCS_ImportObjects = VCS_ImportObjects + 1
                    Case "Module"
                        ' Only import non VCS module files
                        If _
                            VCS_ImportExport.IsNotVCS(objName) _
                        Then
                            VCS_ImportObject objType, objName, objPath
                            VCS_ImportObjects = VCS_ImportObjects + 1
                        Else
                            If VCS_ImportExport.ArchiveMyself Then
                                MsgBox "Module " & objName & " could not be updated while running. Ensure latest version is included!", vbExclamation, "Warning"
                            End If
                        End If
                    Case Else
                        VCS_ImportObject objType, objName, objPath
                        VCS_ImportObjects = VCS_ImportObjects + 1
                End Select

                fileName = Dir$()
            Loop
    End Select
End Function

Public Sub VCS_SaveAccessTypeAsText(ByVal objType As String, ByVal objName As String, ByVal fileName As String)
    On Error GoTo SaveError

    If VCS_ShouldHandleUcs2Conversion(objType) Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        Application.SaveAsText VCS_GetAccessTypeForObjectType(objType), objName, tempFileName
        VCS_File.VCS_ConvertUcs2Utf8 tempFileName, fileName
    Else
        Application.SaveAsText VCS_GetAccessTypeForObjectType(objType), objName, fileName
    End If

    ' Format XML tags of data macro XML file into multiple lines for easier versioning
    If objType = "TableDataMacro" Then
        VCS_FormatXMLTagsIntoMultipleLines fileName
    End If

    Exit Sub

SaveError:
    Select Case Err.Number
        Case 2950
            ' Object does not exist, we just continue in that case
        Case Else
            Debug.Print "ERROR: " & Err.Description & "(" & Err.Number & ")"
    End Select
End Sub

Public Sub VCS_SaveLinkedTableAsText(ByVal objName As String, ByVal fileName As String)
    VCS_ExportLinkedTable objName, fileName
End Sub

Public Sub VCS_SaveTableAsText(ByVal objName As String, ByVal fileName As String)
    Application.ExportXML _
        ObjectType := acExportTable, _
        DataSource := objName, _
        DataTarget := fileName
End Sub

Public Sub VCS_SaveTableDefinitionAsText(ByVal objName As String, ByVal fileName As String)
    Application.ExportXML _
        ObjectType := acExportTable, _
        DataSource := objName, _
        SchemaTarget := fileName
End Sub

Public Sub VCS_SaveReferencesAsText(ByVal fileName As String)
    VCS_ExportReferencesToPath(fileName)
End Sub

Public Sub VCS_SaveRelationAsText(ByVal objName As String, ByVal fileName As String)
    Dim dbSource As Object
    Dim rel As DAO.Relation

    Set dbSource = CurrentDb()
    Set rel = dbSource.Relations(objName)

    VCS_ExportRelation rel, fileName
End Sub

Public Sub VCS_SaveAsText(ByVal objType As String, ByVal objName As String, ByVal fileName As String)
    Select Case objType
        Case "Query", "Form", "Report", "Macro", "Module", "TableDataMacro"
            VCS_SaveAccessTypeAsText objType, objName, fileName
        Case "LinkedTable"
            VCS_SaveLinkedTableAsText objName, fileName
        Case "Table"
            VCS_SaveTableAsText objName, fileName
        Case "TableDefinition"
            VCS_SaveTableDefinitionAsText objName, fileName
        Case "Reference"
            VCS_SaveReferencesAsText fileName
        Case "Relation"
            VCS_SaveRelationAsText objName, fileName
        Case "PrintVars"
            VCS_ExportPrintVars objName, fileName
    End Select
End Sub

Public Sub VCS_ExportObject(ByVal objType As String, ByVal objName As String, ByVal strPath As String)
    Dim strObjectFileName As String
    strObjectFilename = VCS_ObjectNameToPathName(objName) & "." & VCS_GetObjectFileExtension(objType)
    VCS_SaveAsText objType, objName, VCS_AppendDirectoryDelimiter(strPath) & strObjectFilename
End Sub

' Creates directory to store serialized files for object type and removes existing ones
Public Sub VCS_InitializeObjectStorage(ByVal objType As String, ByVal strPath As String)
    VCS_Dir.VCS_MkDirIfNotExist strPath
    VCS_DeleteFilesFromDirByExtension VCS_AppendDirectoryDelimiter(strPath), VCS_GetObjectFileExtension(objType)
End Sub

Public Function VCS_ExportObjects(ByVal objType As String, ByVal strPath As String, Optional ByVal strTablesToExportWithData As String = "*")
    Dim dbSource As Object
    Dim doc As Object
    Dim tdf As DAO.TableDef
    Dim rel As DAO.Relation
    Dim objPath As String

    Set dbSource = CurrentDb()

    VCS_ExportObjects = 0
    objPath = VCS_ObjectPath(strPath, objType)

    ' Init storage
    Select Case objType
        Case "Table"
            VCS_InitializeObjectStorage "LinkedTable", VCS_ObjectPath(strPath, "LinkedTable")
            VCS_InitializeObjectStorage "TableDefinition", VCS_ObjectPath(strPath, "TableDefinition")
            VCS_InitializeObjectStorage "TableDataMacro", VCS_ObjectPath(strPath, "TableDataMacro")
            VCS_InitializeObjectStorage objType, objPath
        Case "Reference"
            ' Single file saved in the root path, thus no need to delete anything
        Case "Report"
            VCS_InitializeObjectStorage "PrintVars", VCS_ObjectPath(strPath, "PrintVars")
            VCS_InitializeObjectStorage objType, objPath
        Case Else
            VCS_InitializeObjectStorage objType, objPath
    End Select

    ' Serialize
    Select Case objType
        Case "Query"
            For Each doc In dbSource.QueryDefs
                ' Skip internal queries
                If Left$(doc.name, 1) <> "~" Then
                    VCS_ExportObject objType, doc.name, objPath
                    VCS_ExportObjects = VCS_ExportObjects + 1
                End If
            Next
        Case "Form"
            For Each doc In dbSource.Containers(VCS_GetContainerNameForObjectType(objType)).Documents
                ' Skip temporary forms
                If Left$(doc.name, 1) <> "~" Then
                    VCS_ExportObject objType, doc.name, objPath
                    VCS_ExportObjects = VCS_ExportObjects + 1
                End If
            Next
        Case "Report"
            For Each doc In dbSource.Containers(VCS_GetContainerNameForObjectType(objType)).Documents
                ' Skip temporary reports
                If Left$(doc.name, 1) <> "~" Then
                    VCS_ExportObject objType, doc.name, objPath
                    VCS_ExportObject "PrintVars", doc.name, VCS_ObjectPath(strPath, "PrintVars")
                    VCS_ExportObjects = VCS_ExportObjects + 1
                End If
            Next
        Case "Macro"
            For Each doc In dbSource.Containers(VCS_GetContainerNameForObjectType(objType)).Documents
                ' Skip temporary macros
                If Left$(doc.name, 1) <> "~" Then
                    VCS_ExportObject objType, doc.name, objPath
                    VCS_ExportObjects = VCS_ExportObjects + 1
                End If
            Next
        Case "Module"
            For Each doc In dbSource.Containers(VCS_GetContainerNameForObjectType(objType)).Documents
                ' Skip temporary modules
                If Left$(doc.name, 1) <> "~" Then
                    ' Skip VCS modules if not wanted
                    If (VCS_ImportExport.IsNotVCS(doc.name) Or VCS_ImportExport.ArchiveMyself) Then
                        VCS_ExportObject objType, doc.name, objPath
                        VCS_ExportObjects = VCS_ExportObjects + 1
                    End If
                End If
            Next
        Case "Reference"
            VCS_ExportObject objType, "references", objPath
        case "Table"
            For Each tdf In dbSource.TableDefs
                ' Skip system, temporary tables and those not to include
                If _
                    Left$(tdf.name, 4) <> "MSys" And _
                    Left$(tdf.name, 1) <> "~" _
                Then
                    If tdf.Connect <> "" Then
                        ' Linked table
                        VCS_ExportObject "LinkedTable", tdf.name, VCS_ObjectPath(strPath, "LinkedTable")
                        VCS_ExportObjects = VCS_ExportObjects + 1
                    Else
                        ' Regular table
                        VCS_ExportObject "TableDefinition", tdf.name, VCS_ObjectPath(strPath, "TableDefinition")
                        VCS_ExportObject "TableDataMacro", tdf.name, VCS_ObjectPath(strPath, "TableDataMacro")
                        If _
                            VCS_HasToken(strTablesToExportWithData, "*") Or _
                            VCS_HasToken(strTablesToExportWithData, tdf.name) _
                        Then
                            VCS_ExportObject objType, tdf.name, objPath
                        End If
                        VCS_ExportObjects = VCS_ExportObjects + 1
                    End If
                End If
            Next
        Case "Relation"
            For Each rel In dbSource.Relations
                ' Exclude relations from system tables and inherited (linked) relations
                ' Skip if dbRelationDontEnforce property is not set. The relationship is already in the table xml file. - sean
                If _
                    Not (rel.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or _
                    rel.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups" Or _
                    (rel.Attributes And DAO.RelationAttributeEnum.dbRelationInherited) = _
                        DAO.RelationAttributeEnum.dbRelationInherited) And _
                    (rel.Attributes = DAO.RelationAttributeEnum.dbRelationDontEnforce) _
                Then
                    VCS_ExportObject objType, rel.name, objPath
                    VCS_ExportObjects = VCS_ExportObjects + 1
                End If
            Next
    End Select
End Function

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

' Sanitizes files for an object type within a directory
Public Sub VCS_SanitizeFilesForObjectTypeAtPath(ByVal objType As String, ByVal strPath As String)
    Select Case objType
        ' Form Specification Syntax
        Case "Query", "Form", "Report", "Macro"
            VCS_SanitizeTextFiles VCS_AppendDirectoryDelimiter(strPath), VCS_GetObjectFileExtension(objType)
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