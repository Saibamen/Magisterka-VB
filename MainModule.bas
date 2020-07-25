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
    FileModule.FilePath = "C:\Users\Adam\source\repos\magisterka\Magisterka-VB\da51f72f-7804-40fe-bc66-8fc5418325fb_001.data"
    
    ' Print empty line
    Debug.Print
    
    ' Use QueryPerformanceCounter instead of the more inaccurate GetTickCount
    Dim stopwatch As Variant
    stopwatch = TimerEx
    
    Dim x As Integer
    For x = 1 To TestAttempts
        'Debug.Print "TestAttempts: " & x
        Call ReadFile_AllText
    Next
    
    stopwatch = TimerEx - stopwatch
    Debug.Print "All tests executed in " & stopwatch & " seconds"
    'MsgBox stopwatch & " seconds", 0, "Main"
End Sub

Public Sub PrintElapsedTime(testName As String, stopwatch As Variant)
   Debug.Print testName & " N = " & Iterations & " = " & stopwatch & " seconds"
   MsgBox stopwatch & " seconds", 0, testName
End Sub
