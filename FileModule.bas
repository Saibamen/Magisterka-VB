Attribute VB_Name = "FileModule"
Option Explicit

Public FilePath As String

Public Function ReadFile_AllText()
    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim returnVar As String
    Dim x As Integer
    For x = 1 To MainModule.Iterations
        'Debug.Print "Iterations: " & x
        returnVar = ReadFileIntoString()
    Next
    
    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("ReadFile_AllText", stopwatch)
End Function

Public Function ReadFileIntoString() As String
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(FilePath)
    ReadFileIntoString = ts.ReadAll
End Function

