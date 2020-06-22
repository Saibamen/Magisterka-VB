Attribute VB_Name = "FileModule"
Option Explicit

Public FilePath As String

Public Function ReadFile_ReadAllText()
    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim Text As String
    Dim x As Integer
    For x = 1 To MainModule.Iterations
        'Debug.Print "Iterations: " & x
        Text = ReadFileIntoString()
    Next
    
    stopwatch = TimerEx - stopwatch
    Debug.Print "ReadFile_ReadAllText: " & stopwatch
    'MsgBox stopwatch & " seconds", 0, "ReadFile_ReadAllText"
End Function

Public Function ReadFileIntoString() As String
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(FilePath)
    ReadFileIntoString = ts.ReadAll
End Function

