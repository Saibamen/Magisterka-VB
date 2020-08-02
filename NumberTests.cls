VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub IntTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = Int(123456789.987654)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("IntTest", stopwatch, testIterations, True)
End Sub
