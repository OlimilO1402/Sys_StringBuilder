VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Beams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mPB As PictureBox
Attribute mPB.VB_VarHelpID = -1
Private mCol As Collection

Private Sub Class_Initialize()
  Clear
End Sub
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub

Public Sub NewC(aPB As PictureBox)
  Set mPB = aPB
End Sub

Public Sub Clear()
  Set mCol = New Collection
End Sub

Public Sub Add(aBeam As Beam)
  Call mCol.Add(aBeam)
End Sub

Public Property Get Beam(Index As Long) As Beam
  Set Beam = mCol.Item(Index)
End Property

Public Property Get Items() As Collection
  Set Items = mCol
End Property
Public Property Get Count() As Long
  Count = mCol.Count
End Property
Public Sub Invalidate()
  Call mPB_Resize
End Sub
Private Sub mPB_Resize()
Dim B As Beam
Dim i As Long
Dim m As Single
Dim L As Single
Dim brdr As Single: brdr = 8 * 15
  m = GetMax
  mPB.Cls
  For i = 0 To mCol.Count - 1
    Set B = mCol.Item(i + 1)
    Call B.Move(mPB.ScaleWidth, mPB.ScaleHeight, brdr, m, i, mCol.Count)
  Next
  'und die Namen anzeigen
  For i = 0 To mCol.Count - 1
    Set B = mCol.Item(i + 1)
    L = brdr + ((mPB.ScaleWidth - brdr) * i / mCol.Count)
    mPB.ForeColor = &HFFFFFF
    mPB.PSet (L, 0)
    mPB.ForeColor = &H0&
    mPB.Print (B.Name)
  Next
End Sub

Private Function GetMax() As Single
Dim B As Beam
  For Each B In mCol
    GetMax = MaxS(GetMax, B.Value)
  Next
End Function
Private Function MaxS(SngVal1 As Single, SngVal2 As Single) As Single
  If SngVal1 > SngVal2 Then MaxS = SngVal1 Else MaxS = SngVal2
End Function
