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

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub IntTest()
    Dim testIterations As Long
    testIterations = 1000000
    Dim returnVar As Integer

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To testIterations
        returnVar = Int(32000.9876545569)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("IntTest", stopwatch, testIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub RoundDecimalPlacesTest()
    Dim testIterations As Long
    testIterations = 1000000
    Dim returnVar As Double

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To testIterations
        returnVar = Format(32000.9876545569, "0.###")
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("RoundDecimalPlacesTest", stopwatch, testIterations)
End Sub
