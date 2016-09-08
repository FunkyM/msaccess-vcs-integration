Attribute VB_Name = "VCS_String"
Option Compare Database

Option Private Module
Option Explicit

' Pad a string on the right to make it `count` characters long.
Public Function VCS_PadRight(ByVal Value As String, ByVal Count As Integer) As String
    VCS_PadRight = Value
    If Len(Value) < Count Then
        VCS_PadRight = VCS_PadRight & Space$(Count - Len(Value))
    End If
End Function

Public Function VCS_HasToken(ByVal search As String, ByVal needle As String, Optional ByVal Delimiter As String = ",") As Boolean
    Dim tokens As Variant
    Dim value As Variant

    ' Remove spaces and split into array
    tokens = Split(Replace(search, " ", vbNullString), Delimiter)

    ' Search for needle
    For Each value In tokens
        If value = needle Then
            VCS_HasToken = True
            Exit Function
        End If
    Next

    VCS_HasToken = False
End Function