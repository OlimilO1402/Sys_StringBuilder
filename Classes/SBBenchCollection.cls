VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BenchCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCol As Collection

Private Sub Class_Initialize()
    Call Clear
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Sub Clear()
    Set mCol = New Collection
End Sub

Public Sub Add(Key As String, aBench As IBench)
    Call mCol.Add(aBench, Key)
End Sub

Public Property Get BenchMark(Key As String) As IBench
    Set BenchMark = mCol.Item(Key)
End Property
