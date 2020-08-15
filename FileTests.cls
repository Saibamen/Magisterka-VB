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
    
    Debug.Print ("Deleting test files...")
    
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

    Dim i As Long
    For i = 1 To MainModule.Iterations
        Dim fileSystemObject As New fileSystemObject
        Dim textStream As textStream
    
        Set textStream = fileSystemObject.OpenTextFile(ReadTestFile, ForReading)
    
        returnVar = textStream.ReadAll()
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("ReadFile_AllText", stopwatch)
End Sub

' NOTE: Decreased to 25 iterations
Public Sub ReadFile_ByLine()
    Dim testIterations As Long
    testIterations = 25
    Dim returnVar As String
    
    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To testIterations
        Dim line As String
        Dim fileNumber As Integer
        fileNumber = FreeFile
        
        Open ReadTestFile For Input As fileNumber
            Do While Not EOF(fileNumber)
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

    Dim i As Long
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

    Dim testFileLines() As String
    testFileLines = ReadTestFileContentAsArray()

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
    For i = 1 To MainModule.Iterations
        Dim filename As String
        filename = TestFilePrefix & i & "_WriteAllLines" & TestFileExtension
        Dim fileNumber As Integer
        fileNumber = FreeFile
        
        Open TestFilesDirectory & "\" & filename For Output As fileNumber
            Dim n As Integer
            For n = 1 To UBound(testFileLines)
                Print #fileNumber, testFileLines(n)
            Next
        Close fileNumber
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("WriteFile_ByLine", stopwatch)
End Sub

Public Sub RenameFiles()
    Call CreateTestFiles

    Dim stopwatch As Variant
    stopwatch = TimerEx

    Dim i As Long
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

    Dim i As Long
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

    Dim i As Long
    For i = 1 To MainModule.Iterations
        Kill TestFilesDirectory & "\" & TestFilePrefix & i & TestFileExtension
    Next

    stopwatch = TimerEx - stopwatch
    Call PrintElapsedTime("DeleteFiles", stopwatch)
End Sub

Private Sub CreateTestFiles()
    Call DeleteTestFiles
    
    Debug.Print ("Creating test files...")
    
    Dim testFileContent As String
    testFileContent = ReadTestFileContent()
    
    ' Create directory if not exist
    If Dir(TestFilesDirectory, vbDirectory) = "" Then
        MkDir TestFilesDirectory
    End If
    
    Dim i As Long
    For i = 1 To MainModule.Iterations
        Dim filename As String
        filename = TestFilePrefix & i & TestFileExtension
        Dim fileNumber As Integer
        fileNumber = FreeFile
        
        Open TestFilesDirectory & "\" & filename For Output As fileNumber
            Print #fileNumber, testFileContent
        Close fileNumber
    Next
    
    Call LogText
End Sub

Private Function ReadTestFileContent() As String
    Dim fileSystemObject As New fileSystemObject
    Dim textStream As textStream

    Set textStream = fileSystemObject.OpenTextFile(ReadTestFile, ForReading)

    ReadTestFileContent = textStream.ReadAll()
    textStream.Close
End Function

Private Function ReadTestFileContentAsArray() As String()
    Dim fileSystemObject As New fileSystemObject
    Dim textStream As textStream
    
    Dim fileArray() As String
    Dim arraySize As Integer
    arraySize = 1

    Set textStream = fileSystemObject.OpenTextFile(ReadTestFile, ForReading)

    Do Until textStream.AtEndOfStream
        ' Resize array and keep existing data
        ReDim Preserve fileArray(arraySize)
        fileArray(arraySize) = textStream.ReadLine
        arraySize = arraySize + 1
    Loop
    
    textStream.Close
    ReadTestFileContentAsArray = fileArray
End Function
