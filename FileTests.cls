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
    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim returnVar As String
    Dim x As Integer
    For x = 1 To MainModule.Iterations
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
