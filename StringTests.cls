VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TestString As String
Public TestStringWithoutSpaces As String

' NOTE: Increased to 30000 iterations
Public Sub AscTest()
    Dim testIterations As Integer
    testIterations = 30000
    Dim returnVar As Integer

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = Asc(TestStringWithoutSpaces)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("AscTest", stopwatch, testIterations, True)
End Sub
