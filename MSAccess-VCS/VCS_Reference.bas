Attribute VB_Name = "VCS_Reference"
Option Compare Database

Option Private Module
Option Explicit

' Imports References from a path
Public Function VCS_ImportReferencesFromPath(ByVal fileName As String) As Integer
    Dim FSO As Object
    Dim InFile As Object
    Dim line As String
    Dim item() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim refName As String

    refName = Dir$(fileName)
    If Len(refName) = 0 Then
        VCS_ImportReferencesFromPath = 0
        Exit Function
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(fileName, iomode:=ForReading, create:=False, Format:=TristateFalse)

    On Error GoTo failed_guid

    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        item = Split(line, ",")
        If UBound(item) = 2 Then 'a ref with a guid
          GUID = Trim$(item(0))
          Major = CLng(item(1))
          Minor = CLng(item(2))
          Application.References.AddFromGuid GUID, Major, Minor
          VCS_ImportReferencesFromPath = VCS_ImportReferencesFromPath + 1
        Else
          refName = Trim$(item(0))
          Application.References.AddFromFile refName
          VCS_ImportReferencesFromPath = VCS_ImportReferencesFromPath + 1
        End If
go_on:
    Loop

    On Error GoTo 0
    InFile.Close
    Set InFile = Nothing
    Set FSO = Nothing
    Exit Function

failed_guid:
    If Err.Number = 32813 Then
        ' The reference is already present in the access project - so we can ignore the error
        Resume Next
    Else
        MsgBox "Failed to register " & GUID, , "Error: " & Err.Number
        ' Do we really want to carry on the import with missing references??? - Surely this is fatal
        Resume go_on
    End If
End Function

' Export References to a path
Public Function VCS_ExportReferencesToPath(ByVal fileName As String) As Integer
    Dim FSO As Object
    Dim OutFile As Object
    Dim line As String
    Dim ref As Reference

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(fileName, overwrite:=True, Unicode:=False)
    VCS_ExportReferencesToPath = 0
    For Each ref In Application.References
        If ref.GUID <> vbNullString Then ' references of types mdb,accdb,mde etc don't have a GUID
            If Not ref.BuiltIn Then
                line = ref.GUID & "," & CStr(ref.Major) & "," & CStr(ref.Minor)
                OutFile.WriteLine line
                VCS_ExportReferencesToPath = VCS_ExportReferencesToPath + 1
            End If
        Else
            line = ref.FullPath
            OutFile.WriteLine line
            VCS_ExportReferencesToPath = VCS_ExportReferencesToPath + 1
        End If
    Next
    OutFile.Close
End Function