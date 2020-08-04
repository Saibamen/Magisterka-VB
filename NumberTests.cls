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

Public NumberTestsIterations As Long

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub IntTest()
    Dim returnVar As Integer

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To NumberTestsIterations
        returnVar = Int(32000.9876545569)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("IntTest", stopwatch, NumberTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub RoundDecimalPlacesTest()
    Dim returnVar As Double

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To NumberTestsIterations
        returnVar = Round(32000.9876545569, 3)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("RoundDecimalPlacesTest", stopwatch, NumberTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub BasicMathTest()
    Dim returnVar As Double

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To NumberTestsIterations
        returnVar = (51564981649.3 - 864518.9 + 9841598198.3 * 11869.4) / 6.7423
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("BasicMathTest", stopwatch, NumberTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations
Public Sub RandomNumberTest()
    Dim returnVar As Integer

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To NumberTestsIterations
        Randomize Timer
        returnVar = Int((2560 - 256 + 1) * Rnd + 256)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("RandomNumberTest", stopwatch, NumberTestsIterations)
End Sub
