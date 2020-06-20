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

Dim dblTimerDauer As Variant

Dim lngTime As Long
Dim lngIndex As Long

dblTimerDauer = TimerEx
lngTime = GetTickCount()

Dim Caption As String

For lngIndex = 1 To 10000
    Caption = CStr(lngIndex)
Next lngIndex

lngTime = GetTickCount - lngTime
MsgBox (TimerEx - dblTimerDauer) & " Sekunden", 0, "Daten einlesen"
Debug.Print "Execution took " & CStr(lngTime); " ms"

End Sub
