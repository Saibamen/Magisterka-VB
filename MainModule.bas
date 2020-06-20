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
TestAttempts = 10
Iterations = 1000
Dim FilePath As String
FilePath = "C:\Users\Saibamen\source\repos\Magisterka\VB6\da51f72f-7804-40fe-bc66-8fc5418325fb_001.data"

Dim dblTimerDauer As Variant

Dim lngTime As Long
Dim lngIndex As Long

' Use QueryPerformanceCounter instead of the more inaccurate GetTickCount
dblTimerDauer = TimerEx
lngTime = GetTickCount()

Dim Caption As String

' Measure
Dim Text As String

Dim x As Integer
For x = 1 To Iterations
    Text = ReadFileIntoString(FilePath)
Next

lngTime = GetTickCount - lngTime
MsgBox (TimerEx - dblTimerDauer) & " Sekunden", 0, "Daten einlesen"
Debug.Print "Execution took " & CStr(lngTime); " ms"
End Sub



Public Function ReadFileIntoString(strFilePath As String) As String
    Dim fso As New FileSystemObject
    Dim ts As TextStream

    Set ts = fso.OpenTextFile(strFilePath)
    ReadFileIntoString = ts.ReadAll
End Function
