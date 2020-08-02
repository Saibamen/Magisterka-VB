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

Public StringTestsIterations As Long
Public TestString As String

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub AscTest()
    Dim returnVar As Integer

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        returnVar = Asc(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("AscTest", stopwatch, StringTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub MidTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        returnVar = Mid(TestString, 20, 50)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("MidTest", stopwatch, StringTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub LeftTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        returnVar = Left(TestString, 50)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("LeftTest", stopwatch, StringTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub RightTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        returnVar = Right(TestString, 50)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("RightTest", stopwatch, StringTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub TrimTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        Dim textToTrim As String
        ' 520 characters, 10 whitespaces both at beginning and end
        textToTrim = "          ThisIsExample123StringToTestStringOperationsTESTweThisIsExample123StringToTestStringOperationsTEST!@ThisIsExample123StringToTestStringOperationsTESTweThisIsExample123StringToTestStringOperationsTEST!@ThisIsExample123StringToTextToSearcherationsTESTweThisIsExample123StringToTestStringOperationsTEST!@ThisIsExample123StringToTestStringOperationsTESTweThisIsExample123StringToTestStringOperationsTEST!@ThisIsExample123StringToTestStringOperationsTESTweThisIsExample123StringToTestStringOperationsTEST!@          "
        returnVar = Trim(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("TrimTest", stopwatch, StringTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations. Time in milliseconds
Public Sub LenTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        returnVar = Len(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("LenTest", stopwatch, StringTestsIterations, True)
End Sub

' NOTE: Increased to 1000000 iterations
Public Sub LCaseTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        returnVar = LCase(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("LCaseTest", stopwatch, StringTestsIterations)
End Sub

' NOTE: Increased to 1000000 iterations
Public Sub UCaseTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        returnVar = UCase(TestString)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("UCaseTest", stopwatch, StringTestsIterations)
End Sub

' NOTE: Increased to 1000000 iterations
Public Sub ReplaceTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        returnVar = Replace(TestString, "TextToSearch", "ChangedText1")
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("ReplaceTest", stopwatch, StringTestsIterations)
End Sub

' NOTE: Increased to 1000000 iterations
Public Sub PadLeftTest()
    Dim returnVar As String

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To StringTestsIterations
        Dim missingWidth As Integer
        ' First number is total width we want to achieve like in C# PadLeft()
        missingWidth = 1000 - Len(TestString)
        returnVar = Right(String(missingWidth, "#") & TestString, Len(TestString) + missingWidth)
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("PadLeftTest", stopwatch, StringTestsIterations)
End Sub
