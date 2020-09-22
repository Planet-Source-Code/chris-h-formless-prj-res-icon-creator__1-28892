Attribute VB_Name = "modFileFolderExists"
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
                                                (ByVal lpFileName As String, _
                                            lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" _
                (ByVal hFindFile As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Function QualifyPath( _
    sSource As String _
) As Boolean
   Dim WFD As WIN32_FIND_DATA
   Dim lFile As Long
   lFile = FindFirstFile(sSource, WFD)
   QualifyPath = lFile <> INVALID_HANDLE_VALUE
   Call FindClose(lFile)
End Function

Public Function QS( _
    sPath As String _
) As String
    If right(sPath, 1) <> "\" Then
        QS = sPath & "\"
    Else
        QS = sPath
    End If
End Function

Public Function ParseFileName( _
    sFileIn As String _
) As String
    Dim i As Integer
    For i = Len(sFileIn) To 1 Step -1
        If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
    Next
    ParseFileName = Mid$(sFileIn, i + 1, Len(sFileIn) - i)
End Function

Public Function ParsePath( _
    sPathIn As String _
) As String
    Dim i As Integer
    For i = Len(sPathIn) To 1 Step -1
        If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
    Next
    ParsePath = left$(sPathIn, i)
End Function

