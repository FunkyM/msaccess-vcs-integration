Option Compare Database
Option Explicit

Public Function DemoHello()
    Debug.Print "Hello world."
End Function

Public Sub LoadVCSModules()
    Application.LoadFromText acModule, "VCS_Loader", CurrentProject.Path & "\..\VCS_Loader.bas"
    Application.Run "loadVCS", CurrentProject.Path & "\..\MSAccess-VCS\"
End Sub

Public Sub UnloadVCSModules()
    Application.Run "unloadVCS"
End Sub

Public Sub ExportSources()
    Application.Run "VCS_ExportAllSources"
End Sub

Public Sub ImportSources()
    Application.Run "VCS_ImportAllSources"
End Sub