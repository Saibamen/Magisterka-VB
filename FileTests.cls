VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BaseDirectory As String
Public TestFilesDirectory As String
Public ReadTestFile As String

Public TestFilePrefix As String
Public TestFileExtension As String

Public Sub DeleteTestFiles()
    ' Create directory if not exist
    If Dir(TestFilesDirectory, vbDirectory) = "" Then
        MkDir TestFilesDirectory
        Exit Sub
    End If
    
    Debug.Print "Deleting test files..."
    
    If Dir(TestFilesDirectory & "\*") <> "" Then
        Kill TestFilesDirectory & "\*"
    End If
    
    If Dir(TestFilesDirectory & "\CopyDirectory\*") <> "" Then
        Kill TestFilesDirectory & "\CopyDirectory\*"
    End If
End Sub

Public Sub ReadFile_AllText()
    Dim returnVar As String
    
    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To MainModule.Iterations
        Dim fileSystemObject As New fileSystemObject
        Dim textStream As textStream
    
        Set textStream = fileSystemObject.OpenTextFile(ReadTestFile)
    
        returnVar = textStream.ReadAll()
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("ReadFile_AllText", stopwatch)
End Sub

' NOTE: Decreased to 25 iterations
Public Sub ReadFile_ByLine()
    Dim testIterations As Integer
    testIterations = 25
    Dim returnVar As String
    
    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To testIterations
        Dim line As String
        Dim fileNumber As Integer
        fileNumber = FreeFile
        
        Open ReadTestFile For Input As fileNumber
            Do While Not EOF(1)
                Line Input #fileNumber, line
                returnVar = returnVar + line
            Loop
        Close fileNumber
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("ReadFile_ByLine", stopwatch, testIterations)
End Sub

Public Sub WriteFile_AllText()
    Call DeleteTestFiles

    Dim testFileContent As String
    testFileContent = ReadTestFileContent()

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To MainModule.Iterations
        Dim filename As String
        filename = TestFilePrefix & i & "_WriteAllText" & TestFileExtension
        Dim fileNumber As Integer
        fileNumber = FreeFile
        
        Open TestFilesDirectory & "\" & filename For Output As fileNumber
            Print #fileNumber, testFileContent
        Close fileNumber
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("WriteFile_AllText", stopwatch)
End Sub

Public Sub WriteFile_ByLine()
    Call DeleteTestFiles

    Dim testFileContent As String
    testFileContent = ReadTestFileContent()

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To MainModule.Iterations
        ' TODO WriteFile_ByLine
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("WriteFile_ByLine", stopwatch)
End Sub

Public Sub RenameFiles()
    Call CreateTestFiles

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To MainModule.Iterations
        Dim filename As String
        Dim newFilename As String
        filename = TestFilesDirectory & "\" & TestFilePrefix & i & TestFileExtension
        newFilename = TestFilesDirectory & "\" & TestFilePrefix & i & "_renamed" & TestFileExtension
        
        Name filename As newFilename
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("RenameFiles", stopwatch)
End Sub

Public Sub CopyFiles()
    Call CreateTestFiles

    Dim copyDirectory As String
    copyDirectory = TestFilesDirectory & "\CopyDirectory"
    ' Create directory if not exist
    If Dir(copyDirectory, vbDirectory) = "" Then
        MkDir copyDirectory
    End If

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To MainModule.Iterations
        FileCopy TestFilesDirectory & "\" & TestFilePrefix & i & TestFileExtension, copyDirectory & "\" & TestFilePrefix & i & "_copied" & TestFileExtension
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("CopyFiles", stopwatch)
End Sub

Public Sub DeleteFiles()
    Call CreateTestFiles

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Integer
    For i = 1 To MainModule.Iterations
        Kill TestFilesDirectory & "\" & TestFilePrefix & i & TestFileExtension
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("DeleteFiles", stopwatch)
End Sub

Private Sub CreateTestFiles()
    Call DeleteTestFiles
    
    Debug.Print "Creating test files..."
    
    Dim testFileContent As String
    testFileContent = ReadTestFileContent()
    
    ' Create directory if not exist
    If Dir(TestFilesDirectory, vbDirectory) = "" Then
        MkDir TestFilesDirectory
    End If
    
    Dim i As Integer
    For i = 1 To MainModule.Iterations
        Dim filename As String
        filename = TestFilePrefix & i & TestFileExtension
        Dim fileNumber As Integer
        fileNumber = FreeFile
        
        Open TestFilesDirectory & "\" & filename For Output As fileNumber
            Print #fileNumber, testFileContent
        Close fileNumber
    Next
    
    Debug.Print vbNewLine
End Sub

Private Function ReadTestFileContent() As String
    Dim fileSystemObject As New fileSystemObject
    Dim textStream As textStream

    Set textStream = fileSystemObject.OpenTextFile(ReadTestFile)

    ReadTestFileContent = textStream.ReadAll()
End Function
