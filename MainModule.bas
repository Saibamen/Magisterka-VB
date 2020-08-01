Attribute VB_Name = "MainModule"
Option Explicit

Public Iterations As Integer
Private TestRuns As Integer

Private BaseDirectory As String
Private LogFilename As String

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
    Iterations = 1000
    TestRuns = 10
    ' Change path for your user
    BaseDirectory = "C:\Users\Adam\source\repos\magisterka\Magisterka-VB\"
    LogFilename = "TestsOutputVB6.log"
    
    ' Delete previous log file if exist
    If Dir(BaseDirectory & LogFilename) <> "" Then
        Kill BaseDirectory & LogFilename
    End If
    
    Debug.Print
    
    ' Use QueryPerformanceCounter instead of the more inaccurate GetTickCount
    Dim stopwatch As Variant
    stopwatch = TimerEx
    
    ' FileTests
    Call LogText("FileTests" & vbNewLine)
    
    FileTests.TestFilesDirectory = BaseDirectory & "TestFiles"
    FileTests.ReadTestFile = BaseDirectory & "da51f72f-7804-40fe-bc66-8fc5418325fb_001.data"
    
    FileTests.TestFilePrefix = "testFile_"
    FileTests.TestFileExtension = ".txt"
    
    Call CallByName(FileTests, "DeleteTestFiles", VbMethod)

    Call RunTestsFor(FileTests, "ReadFile_AllText")
    Call RunTestsFor(FileTests, "ReadFile_ByLine")
    Call RunTestsFor(FileTests, "WriteFile_AllText")
    Call RunTestsFor(FileTests, "WriteFile_ByLine")
    Call RunTestsFor(FileTests, "RenameFiles")
    Call RunTestsFor(FileTests, "CopyFiles")
    Call RunTestsFor(FileTests, "DeleteFiles")
    
    Call LogText
    
    ' StringTests
    'Call LogText("StringTests" & vbNewLine)

    '
    
    'Call LogText()
    
    ' NumberTests
    'Call LogText("NumberTests" & vbNewLine)

    '
    
    'Call LogText()
    
    stopwatch = TimerEx - stopwatch
    Call LogText("All tests executed in " & stopwatch & " seconds")
    Debug.Print "Log file saved in " & BaseDirectory & LogFilename
End Sub

Public Sub PrintElapsedTime(testName As String, stopwatch As Variant, Optional testIterations As Integer)
    If testIterations = 0 Then
        testIterations = Iterations
    End If

    Call LogText(testName & " N = " & testIterations & " = " & stopwatch & " seconds")
End Sub

Private Sub RunTestsFor(staticClass As Variant, functionName As String)
    Dim i As Integer
    For i = 1 To TestRuns
        Call CallByName(staticClass, functionName, VbMethod)
    Next
    
    Call LogText
End Sub

Public Sub LogText(Optional text As String)
    Debug.Print text
    
    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    Open BaseDirectory & LogFilename For Append As fileNumber
        Print #fileNumber, text
    Close fileNumber
End Sub
