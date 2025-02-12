Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_USER = &H400
Private Const EM_DISPLAYBAND = WM_USER + 51
Private Const cm_CopyFullNamesToClip = 2018

Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function PathCombine Lib "shlwapi.dll" Alias "PathCombineA" (ByVal szDest As String, ByVal lpszDir As String, ByVal lpszFile As String) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long


'------------------------------------------
Public TChwnd As Long
Public INIFile As String
Public DefaultEditor As String
Public LimitFileSize As Long
Public LimitFileCount As Long
Type Editor
    name As String
    path As String
    ext As String
End Type
Public Editors() As Editor
Public Files() As String, FileCount As Long

Sub Main()
'get Total Commander's handle
    TChwnd = FindWindow("TTOTAL_CMD", vbNullString)
        If TChwnd = 0 Then MsgBox "Cannot handle Total Commander": End
    
    Dim FileList As String
    Call ClipboardSave(frmTCF4.hWnd)  '=========================================== SAVE CLIPBOARD
'get selected files with TC's internal cm_CopyFullNamesToClip command
    SendMessage TChwnd, EM_DISPLAYBAND, cm_CopyFullNamesToClip, 0   'copy selected files path+name to clipboard
        'On Local Error Resume Next
    ReDim Editors(0) As Editor
    ReDim Files(0) As String
    
    Dim lStart As Long
    lStart = GetTickCount
    Do While FileList = ""
        DoEvents
        If GetTickCount > (lStart + 5000) Then Exit Do 'break if passed 5 seconds
        Sleep 100
        FileList = Clipboard.GetText
    Loop
        'On Local Error GoTo 0
        'If Err.Number = 521 Then MsgBox "Cannot open Clipboard. Another app. may have it open": End
    If FileList = "" Then MsgBox "Cannot get selected files": End
    
    If InStr(Clipboard.GetText, Trim(Replace(Command, Chr(34), ""))) = 0 Then
        'new file created with Shift+F4
        Files = Split(Command, vbCrLf)
    Else
        'active file or selected files
        Files = Split(FileList, vbCrLf)
    End If
    Call ClipboardRestore   '=========================================== RESTORE CLIPBOARD
    'Clipboard.Clear   '=========================================== Clear CLIPBOARD
'Set INI file. if not exist, create default INI
    INIFile = App.path & "\" & App.exeName & ".ini"
    If Dir(INIFile) = "" Then CreateDefaultINI
    DefaultEditor = getAbsolutePath(INIRead("Editors", "Default"))
    
    LoadINI
    If UBound(Editors) = 0 And DefaultEditor = "" Then
        MsgBox "Undefined default editor/editors" & vbCrLf & "Using notepad.exe"
        DefaultEditor = "notepad.exe"
    End If
    
'process selected files
    Dim FileExt As String, i As Long, j As Long, editorPath As String, passFile As Boolean
    Dim hMenu As Long
    Dim foundEditor As Integer
    FileCount = UBound(Files)
    If FileCount > LimitFileCount - 1 Then
        FileCount = LimitFileCount - 1
        MsgBox "Files passed. File count over FileCount." & vbCrLf & vbCrLf & Files(i)
    End If
    For i = 0 To FileCount
        Files(i) = Replace(Files(i), Chr(34) & Chr(34), Chr(34))
        foundEditor = -1
        FileExt = ""
        passFile = False
        passFile = CBool(Right(Files(i), 1) = "\")   'if not DIRECTORY
        If LimitFileSize > 0 Then
            Dim Pointer As Long, sizeofFile As Long, lpFSHigh As Long
            Pointer = lOpen(Files(i), &H0&)
            sizeofFile = GetFileSize(Pointer, lpFSHigh)
            lclose Pointer
            
            passFile = CBool(sizeofFile > LimitFileSize)
            If passFile Then MsgBox "File passed. Size over LimitFileSize." & vbCrLf & vbCrLf & Files(i)
        End If
        If Not passFile Then
            FileExt = "," & getFileExt(Files(i)) & ","
            For j = 0 To UBound(Editors)
                If InStr(Editors(j).ext, FileExt) > 0 Then
                    foundEditor = j
                    Exit For
                End If
            Next j
            If foundEditor = -1 Then
                'extension not found. Using default editor
                editorPath = getAbsolutePath(DefaultEditor)
            Else
                editorPath = getAbsolutePath(Editors(foundEditor).path)
            End If
            If editorPath <> "" Then
                If Left(Files(i), 1) <> Chr(34) And InStr(Files(i), " ") > 0 Then Files(i) = Chr(34) & Files(i) & Chr(34)
                Shell editorPath & " " & Files(i), vbNormalFocus
            End If
        End If
    Next i
    End
End Sub

Private Function getAbsolutePath(FileName As String) As String
    Dim FullName As String, ret As Long, strOutput As String
    FullName = Trim(FileName)
    strOutput = Space(256)
    If InStr(FileName, "%") Then
        ret = ExpandEnvironmentStrings(FullName, strOutput, Len(strOutput))
        If ret > 0 Then FullName = Left(strOutput, ret - 1)
    ElseIf Left(FullName, 1) = "\" Or Left(FullName, 2) = ".\" Or Left(FullName, 3) = "..\" Then
        If PathCombine(strOutput, App.path, FullName) Then FullName = Left$(strOutput, InStr(strOutput, vbNullChar) - 1)
    End If
    getAbsolutePath = FullName
End Function
Private Function getFileExt(FileName As String) As String
    If InStrRev(FileName, ".") > 0 Then getFileExt = Mid(FileName, InStrRev(FileName, ".") + 1)
    getFileExt = Replace(getFileExt, Chr(34), "")
End Function
Private Function getFileName(FileName As String) As String
    If InStrRev(FileName, ".") > 0 Then getFileName = Left(FileName, InStrRev(FileName, ".") - 1)
    If InStrRev(getFileName, "\") > 0 Then getFileName = Mid(getFileName, InStrRev(getFileName, "\") + 1)
End Function

Sub LoadINI()
    Dim i As Long, exeName As String, exePath As String, exeExt As String
    LimitFileSize = CLng(INIRead("General", "LimitFileSize")): If LimitFileSize <= 0 Then LimitFileSize = 100000000
    LimitFileCount = CLng(INIRead("General", "LimitFileCount")): If LimitFileCount < 1 Then LimitFileCount = 50
    Do
        i = i + 1
        exePath = INIRead("Editors", "Path" & i):   If exePath = "" Then Exit Do
        exeName = INIRead("Editors", "Name" & i)
        exeExt = INIRead("Editors", "Ext" & i): exeExt = Replace(exeExt, " ", ""): If exeExt = "" Then Exit Do
        exeExt = "," & exeExt & ","
        Do While InStr(exeExt, ",,")
            exeExt = Replace(exeExt, ",,", ",")
        Loop
        If exeName = "" Then exeName = getFileName(exePath)
        ReDim Preserve Editors(i - 1) As Editor
        If InStr(exePath, Chr(34)) = 0 And InStr(exePath, " ") > 0 Then exePath = Chr(34) & exePath & Chr(34)
        Editors(i - 1).path = getAbsolutePath(exePath)
        Editors(i - 1).name = exeName
        Editors(i - 1).ext = exeExt
    Loop
End Sub

Sub CreateDefaultINI()
    Dim ff: ff = FreeFile
    Open INIFile For Output As #ff
        Print #ff, "[General]"
            Print #ff, "LimitFileSize=100000000"
            Print #ff, "LimitFileCount=10"
        Print #ff, "[Editors]"
            Print #ff, "Default=%COMMANDER_PATH%\Soft\Notepad++\notepad++.exe -nosession -multiInst"
    Close #ff
End Sub
