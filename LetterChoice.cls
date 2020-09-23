VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LetterChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Letter As String
Public DifferenceValue As Long
Public FileData As String
Public Function PixelValue(Sector As Long) As Long
    Dim tmpFD() As String
    Let tmpFD = Split(Me.FileData, ",")
    Let PixelValue = tmpFD(Sector)
End Function
