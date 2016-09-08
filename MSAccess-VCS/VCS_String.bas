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
