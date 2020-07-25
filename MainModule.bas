Attribute VB_Name = "MainModule"
Option Explicit

Public TestAttempts As Integer
Public Iterations As Integer

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
  lpFrequency As Currency) As Long
 
Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
  lpPerformanceCount As Currency) As Long

Public Function TimerEx() As Currency
    Static nFreq As Currency

    If nFreq = 0 Then
        QueryPerformanceFrequency nFreq
    End If

    Dim nTimer As Currency
    QueryPerformanceCounter nTimer
    TimerEx = nTimer / nFreq
End Function

Sub Main()
    ' Zmiana nazwy plikow
    'Dim FileName As String
    'Dim NewFileName As String
    'FileName = "C:\test\before.txt"
    'NewFileName = "C:\test\after.txt"
    'Name FileName As NewFileName
    
    ' Kopiowanie plikow
    'FileCopy "C:\test\before.txt", "C:\test\after.txt"
    
    ' Usuwanie plikow
    'Kill "c:\doomed_dir\*.*"
    
    TestAttempts = 10
    Iterations = 1000
    ' Change path for your user
    FileTests.FilePath = "C:\Users\Adam\source\repos\magisterka\Magisterka-VB\da51f72f-7804-40fe-bc66-8fc5418325fb_001.data"
    
    Debug.Print vbNewLine
    
    ' Use QueryPerformanceCounter instead of the more inaccurate GetTickCount
    Dim stopwatch As Variant
    stopwatch = TimerEx
    
    ' FileTests
    Debug.Print "FileTests" & vbNewLine

    Call RunTestsFor(FileTests, "ReadFile_AllText")
    Call RunTestsFor(FileTests, "ReadFile_ByLine")
    
    Debug.Print vbNewLine
    
    ' StringTests
    'Debug.Print "StringTests" & vbNewLine

    '
    
    'Debug.Print vbNewLine
    
    ' NumberTests
    'Debug.Print "NumberTests" & vbNewLine

    '
    
    'Debug.Print vbNewLine
    
    stopwatch = TimerEx - stopwatch
    Debug.Print "All tests executed in " & stopwatch & " seconds"
    'MsgBox stopwatch & " seconds", 0, "Main"
End Sub

Public Sub PrintElapsedTime(testName As String, stopwatch As Variant, Optional testIterations As Integer)
    If testIterations = 0 Then
        testIterations = Iterations
    End If

    Debug.Print testName & " N = " & testIterations & " = " & stopwatch & " seconds"
    'MsgBox stopwatch & " seconds", 0, testName
End Sub

Public Sub RunTestsFor(staticClass As Variant, functionName As String)
    Dim x As Integer
    For x = 1 To TestAttempts
        Call CallByName(staticClass, functionName, VbMethod)
    Next
    
    Debug.Print vbNewLine
End Sub
