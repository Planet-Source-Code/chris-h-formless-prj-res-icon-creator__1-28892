VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Enum eFileAccess
    eFileRead = &H80000000
    eFileWrite = &H40000000
End Enum

Public Enum eFileCommand
    eFileCreate = 1
    eFileCreatePlus = 2
    eFileOpen = 3
    eFileOpenPlus = 4
    eFileTruncate = 5
End Enum

Public Enum eFilePos
    eFileBegin = 0
    eFileCurrent = 1
    eFileEnd = 2
End Enum

Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Dim hFile As Long
Dim sFileName As String

Public Function OpenAPI(sFile As String, Optional ByVal eAccess As eFileAccess = eFileRead + eFileWrite, _
    Optional ByVal eCommand As eFileCommand = eFileOpenPlus) As Long
sFileName = sFile
hFile = CreateFile(sFile & vbNullChar, eAccess, FILE_SHARE_READ, 0, eCommand, FILE_ATTRIBUTE_NORMAL + FILE_FLAG_SEQUENTIAL_SCAN, 0)
OpenAPI = hFile
End Function

Public Sub CloseAPI()
If hFile <> 0 Then
    CloseHandle hFile
End If
End Sub

Public Function ReadAPI(Text As String, Optional ByVal lReadLength As Long = 0) As Long
Dim lActualLength As Long, Ret As Long
Dim byteData() As Byte

ReadAPI = 0
If hFile = 0 Then Exit Function

If lReadLength = 0 Then
    lReadLength = FileSize
    FilePos = 0
End If
If FilePos + lReadLength > FileSize Then
    lReadLength = FileSize - FilePos
End If

ReDim byteData(lReadLength)

Ret = ReadFile(hFile, byteData(0), lReadLength, lActualLength, 0)

If Ret = 0 Then Exit Function

If lActualLength <> 0 Then
    Text = Space(lActualLength)
    Text = StrConv(byteData, vbUnicode)
    Text = left(Text, lActualLength)
Else
    Text = ""
End If
ReadAPI = lActualLength
End Function

Public Function WriteAPI(Text As String) As Long
Dim TextLength As Long, WriteLength As Long, Ret As Long, i As Long
Dim byteData() As Byte

WriteAPI = 0
If hFile = 0 Then Exit Function

TextLength = Len(Text)
ReDim byteData(TextLength)
byteData() = StrConv(Text & vbNullString, vbFromUnicode)

Ret = WriteFile(hFile, byteData(0), TextLength, WriteLength, ByVal 0)

If Ret = 0 Then Exit Function
WriteAPI = WriteLength
End Function

Public Function AppendAPI(Text As String, Optional ByVal ePosition As eFilePos = eFileEnd) As Long
Dim sData As String, lPos As Long, lSize As Long

AppendAPI = 0
If hFile = 0 Then Exit Function

If ePosition = eFileEnd Then
    FilePosEx(eFileEnd) = 0
    AppendAPI = WriteAPI(Text)
Else
    If ePosition = eFileBegin Then
        FilePos = 0
        lPos = 0
    Else
        lPos = FilePos
    End If
    lSize = FileSize
    If lPos = lSize Then
        AppendAPI = WriteAPI(Text)
        Exit Function
    End If
    ReadAPI sData, lSize - lPos
    sData = Text & sData
    FilePos = lPos
    AppendAPI = WriteAPI(sData)
End If
End Function

Public Property Get FileSize() As Long
If hFile <> 0 Then
    FileSize = GetFileSize(hFile, ByVal 0)
Else
    FileSize = 0
End If
End Property

Public Property Let FileSize(ByVal vNewValue As Long)
If hFile <> 0 Then
    SetFilePointer hFile, vNewValue, 0, eFileBegin
    SetEndOfFile hFile
End If
End Property

Public Property Get FilePos() As Long
If hFile <> 0 Then
    FilePos = SetFilePointer(hFile, 0, 0, eFileCurrent)
Else
    FilePos = 0
End If
End Property

Public Property Let FilePos(ByVal vNewValue As Long)
If hFile <> 0 Then
    If vNewValue > FileSize Then vNewValue = FileSize
    SetFilePointer hFile, vNewValue, 0, eFileBegin
End If
End Property

Public Property Let FilePosEx(ByVal ePosition As eFilePos, ByVal vNewValue As Long)
If hFile <> 0 Then
    SetFilePointer hFile, vNewValue, 0, ePosition
End If
End Property

Public Property Get FileHandle() As Long
FileHandle = hFile
End Property

Private Property Let FileHandle(ByVal vNewValue As Long)
hFile = vNewValue
End Property

Public Property Get FileName() As String
FileName = sFileName
End Property

Private Property Let FileName(ByVal vNewValue As String)
sFileName = vNewValue
End Property

Private Sub Class_Terminate()
If hFile <> 0 Then CloseAPI
End Sub
