Attribute VB_Name = "CoreUtilsLibrary"
Option Explicit

' --- Windows API Declarations ---
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

' -------------------------------------------------------------------------------------
' [ UI Helpers ]
' -------------------------------------------------------------------------------------

' Displays a Yes/No question dialog and returns True if 'Yes' is selected
Public Function PromptUserQuestion(ByVal message As String, Optional ByVal defaultToYes As Boolean = True) As Boolean
    Dim config As VbMsgBoxStyle
    config = vbQuestion + vbYesNo + IIf(defaultToYes, vbDefaultButton1, vbDefaultButton2)

    If MsgBox(message, config, "Action Required") = vbYes Then
        PromptUserQuestion = True
    Else
        PromptUserQuestion = False
    End If
End Function

' Opens a standard File Open dialog
Public Function GetOpenFilePath(ByVal filterLabel As String, ByVal filterMask As String) As String
    Dim selectedFile As Variant
    selectedFile = Application.GetOpenFilename(filterLabel & " (" & filterMask & ")," & filterMask)

    If selectedFile <> False Then
        GetOpenFilePath = CStr(selectedFile)
    Else
        GetOpenFilePath = ""
    End If
End Function

' -------------------------------------------------------------------------------------
' [ File System Operations ]
' -------------------------------------------------------------------------------------

' Checks if a file or folder exists at the specified path
Public Function PathExists(ByVal fullPath As String) As Boolean
    If Dir(fullPath, vbNormal + vbDirectory) <> "" Then
        PathExists = True
    Else
        PathExists = False
    End If
End Function

' Extracts the file name from a full path
Public Function GetFileName(ByVal fullPath As String, Optional ByVal removeExtension As Boolean = False) As String
    Dim fileName As String
    fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)

    If removeExtension And InStr(fileName, ".") > 0 Then
        GetFileName = Left(fileName, InStrRev(fileName, ".") - 1)
    Else
        GetFileName = fileName
    End If
End Function

' Returns a list of files in a directory matching a specific mask
Public Function ListFiles(ByVal directoryPath As String, ByVal fileMask As String, ByRef outFiles() As String) As Boolean
    Dim fileName As String
    Dim count As Long: count = 0

    If Right(directoryPath, 1) <> "\" Then directoryPath = directoryPath & "\"

    fileName = Dir(directoryPath & fileMask)
    Do While fileName <> ""
        ReDim Preserve outFiles(count)
        outFiles(count) = fileName
        count = count + 1
        fileName = Dir
    Loop

    ListFiles = (count > 0)
End Function

' -------------------------------------------------------------------------------------
' [ System Info ]
' -------------------------------------------------------------------------------------

' Retrieves the current Windows username using API
Public Function GetCurrentWindowsUser() As String
    Dim buffer As String * 255
    Dim bufferLen As Long: bufferLen = 255

    If WNetGetUser("", buffer, bufferLen) = 0 Then
        GetCurrentWindowsUser = Trim(Left(buffer, InStr(buffer, vbNullChar) - 1))
    Else
        GetCurrentWindowsUser = "Unknown"
    End If
End Function

' -------------------------------------------------------------------------------------
' [ String Manipulation ]
' -------------------------------------------------------------------------------------

' Enhanced Split function that supports multiple delimiters
Public Function SplitExtended(ByVal text As String, ByVal primarySep As String, Optional ByVal secondarySep As String = "") As Variant
    Dim result() As String
    text = Trim(text)

    ' Basic logic for multi-delimiter split
    If secondarySep <> "" And InStr(text, primarySep) = 0 Then
        result = Split(text, secondarySep)
    Else
        result = Split(text, primarySep)
    End If

    Dim i As Long
    for i = LBound(result) to UBound(result)
        result(i) = Trim(result(i))
    Next i

    SplitExtended = result
End Function