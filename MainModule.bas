Attribute VB_Name = "MainModule"
Option Explicit

Public TestAttempts As Integer
Public Iterations As Integer

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
  lpFrequency As Currency) As Long
 
Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
  lpPerformanceCount As Currency) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

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
    FileModule.FilePath = "C:\Users\Saibamen\source\repos\Magisterka\VB6\da51f72f-7804-40fe-bc66-8fc5418325fb_001.data"
    
    ' Print empty line
    Debug.Print
    
    'Dim stopwatch As Variant
    ' Not useGetTickCount
    'Dim lngTime As Long
    
    ' Use QueryPerformanceCounter instead of the more inaccurate GetTickCount
    'stopwatch = TimerEx
    ' Not useGetTickCount
    'lngTime = GetTickCount()
    
    ' Measure
    
    Dim x As Integer
    For x = 1 To TestAttempts
        'Debug.Print "TestAttempts: " & x
        Call ReadFile_AllText
    Next
    
    'lngTime = GetTickCount - lngTime
    'stopwatch = TimerEx - stopwatch
    'Debug.Print "Main: " & stopwatch
    'MsgBox stopwatch & " seconds", 0, "Main"
    ' Not useGetTickCount
    'Debug.Print "Execution took " & CStr(lngTime); " ms"
End Sub

