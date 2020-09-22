Attribute VB_Name = "modRes"
Option Explicit

Public Function LoadDataIntoFile( _
    DataName As Integer, _
    FileName As String _
) As Boolean
    On Error GoTo ER
    
    Dim myArray() As Byte
    Dim myFile As Long
    If Dir(FileName) = "" Then
        myArray = LoadResData(DataName, "CUSTOM")
        myFile = FreeFile
        Open FileName For Binary Access Write As #myFile
        Put #myFile, , myArray
        Close #myFile
    End If
    LoadDataIntoFile = True
    Exit Function
ER:
    LoadDataIntoFile = False
End Function

