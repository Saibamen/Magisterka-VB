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

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub AscTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As Integer

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = Asc(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("AscTest", stopwatch, testIterations, True)
End Sub

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub MidTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = Mid(TestString, 20, 50)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("MidTest", stopwatch, testIterations, True)
End Sub

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub LeftTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = Left(TestString, 50)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("LeftTest", stopwatch, testIterations, True)
End Sub

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub RightTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = Right(TestString, 50)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("RightTest", stopwatch, testIterations, True)
End Sub

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub TrimTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        Dim textToTrim As String
        ' 520 characters, 10 whitespaces at beginning and end
        textToTrim = "          ThisIsExample123StringToTestStringOperationsTESTweThisIsExample123StringToTestStringOperationsTEST!@ThisIsExample123StringToTestStringOperationsTESTweThisIsExample123StringToTestStringOperationsTEST!@ThisIsExample123StringToTextToSearcherationsTESTweThisIsExample123StringToTestStringOperationsTEST!@ThisIsExample123StringToTestStringOperationsTESTweThisIsExample123StringToTestStringOperationsTEST!@ThisIsExample123StringToTestStringOperationsTESTweThisIsExample123StringToTestStringOperationsTEST!@          "
        returnVar = Trim(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("TrimTest", stopwatch, testIterations, True)
End Sub

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub LenTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = Len(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("LenTest", stopwatch, testIterations, True)
End Sub

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub LCaseTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = LCase(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("LCaseTest", stopwatch, testIterations, True)
End Sub

' NOTE: Increased to 32000 iterations. Time in milliseconds
Public Sub UCaseTest()
    Dim testIterations As Integer
    testIterations = 32000
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        returnVar = UCase(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("UCaseTest", stopwatch, testIterations, True)
End Sub
