VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FilePath As String

Public Function ReadFile_AllText()
    Dim returnVar As String
    
    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim x As Integer
    For x = 1 To MainModule.Iterations
        Dim fileSystemObject As New fileSystemObject
        Dim textStream As textStream
    
        Set textStream = fileSystemObject.OpenTextFile(FilePath)
    
        returnVar = textStream.ReadAll()
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("ReadFile_AllText", stopwatch)
End Function

' NOTE: Decreased to 25 iterations
Public Function ReadFile_ByLine()
    Dim testIterations As Integer
    testIterations = 25
    Dim returnVar As String
    
    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim x As Integer
    For x = 1 To testIterations
        Dim line As String
        
        Open FilePath For Input As #1
        Do While Not EOF(1)
            Line Input #1, line
            returnVar = returnVar + line
        Loop
        Close #1
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("ReadFile_ByLine", stopwatch, testIterations)
End Function
