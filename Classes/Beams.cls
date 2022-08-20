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
Private m_Col As Collection
Private m_Inv As Boolean

Private Sub Class_Initialize()
    Clear
End Sub
Private Sub Class_Terminate()
    Set m_Col = Nothing
End Sub

Friend Sub New_(aPB As PictureBox)
    Set mPB = aPB
End Sub

Public Sub Clear()
    Set m_Col = New Collection
End Sub

Public Sub Add(aBeam As Beam)
    m_Col.Add aBeam
End Sub

Public Property Get Beam(Index As Long) As Beam
    Set Beam = m_Col.Item(Index)
End Property

Public Property Get Items() As Collection
    Set Items = m_Col
End Property
Public Property Get Count() As Long
    Count = m_Col.Count
End Property
Public Sub Invalidate()
    Call mPB_Resize
End Sub

Public Property Let Inverted(Value As Boolean)
    m_Inv = Value
End Property
Public Property Get Inverted() As Boolean
    Inverted = m_Inv
End Property

Private Sub mPB_Resize()
    Dim brdr As Single: brdr = 8 * 15
    Dim m As Single
    If m_Inv Then
        m = GetMin
    Else
        m = GetMax
    End If
    'Debug.Print CStr(m_Inv) & "   " & CStr(m)
    mPB.Cls
    Dim B As Beam
    Dim i As Long
    For i = 0 To m_Col.Count - 1
        Set B = m_Col.Item(i + 1)
        B.Move mPB.ScaleWidth, mPB.ScaleHeight, brdr, m, i, m_Col.Count, m_Inv
    Next
    'und die Namen anzeigen
    Dim L As Single
    For i = 0 To m_Col.Count - 1
        Set B = m_Col.Item(i + 1)
        L = brdr + ((mPB.ScaleWidth - brdr) * i / m_Col.Count)
        mPB.ForeColor = &HFFFFFF
        mPB.PSet (L, 0)
        mPB.ForeColor = &H0&
        mPB.Print (B.Name)
    Next
End Sub

Private Function GetMin() As Single
    Dim B As Beam
    GetMin = 2147483647
    For Each B In m_Col
        GetMin = Min(GetMin, B.GetValue(m_Inv))
    Next
End Function

Private Function GetMax() As Single
    Dim B As Beam
    For Each B In m_Col
        GetMax = Max(GetMax, B.GetValue(m_Inv))
    Next
End Function

Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Private Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

