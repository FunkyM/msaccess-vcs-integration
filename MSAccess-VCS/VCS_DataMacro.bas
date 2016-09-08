Attribute VB_Name = "VCS_DataMacro"
Option Compare Database

Option Private Module
Option Explicit

' Splits exported DataMacro XML into multiple lines
' This allows a VCS to find changes within lines using diff
Public Sub VCS_FormatXMLDataMacroFile(ByVal filePath As String)
    Dim saveStream As Object 'ADODB.Stream

    Set saveStream = CreateObject("ADODB.Stream")
    saveStream.Charset = "utf-8"
    saveStream.Type = 2 'adTypeText
    saveStream.Open

    Dim objStream As Object 'ADODB.Stream
    Dim strData As String
    Set objStream = CreateObject("ADODB.Stream")

    objStream.Charset = "utf-8"
    objStream.Type = 2 'adTypeText
    objStream.Open
    objStream.LoadFromFile (filePath)

    Do While Not objStream.EOS
        strData = objStream.ReadText(-2) 'adReadLine

        Dim tag As Variant
        For Each tag In Split(strData, ">")
            If tag <> vbNullString Then
                saveStream.WriteText tag & ">", 1 'adWriteLine
            End If
        Next
    Loop

    objStream.Close
    saveStream.SaveToFile filePath, 2 'adSaveCreateOverWrite
    saveStream.Close
End Sub