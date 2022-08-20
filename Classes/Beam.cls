VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Beam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PB     As PictureBox
Private m_SngVal As Single
Private m_Name   As String
Private m_Color  As Long

Private Sub Class_Initialize()
    'm_SngVal = 1 '2147483647
    m_SngVal = -1# '999999 '2147483647
End Sub

Friend Sub New_(aFrm As Form, aParent As PictureBox, aName As String, Index As Long, Optional aColor As Long = &HFF&)
    Set m_PB = aFrm.Controls.Add("VB.PictureBox", aName & CStr(Index), aParent)
    m_Name = aName
    m_PB.Visible = True
    m_PB.BackColor = aColor
    m_PB.BorderStyle = 0
End Sub

Public Property Let Name(aName As String)
    m_Name = aName
End Property
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Color(aColor As String)
    m_Color = aColor
End Property
Public Property Get Color() As String
    Color = m_Color
End Property

'Public Property Get PB() As PictureBox
'  Set PB = m_PB
'End Property

Public Property Let Value(aSngVal As Single)
    m_SngVal = aSngVal
End Property
Public Property Get Value() As Single
    Value = m_SngVal
End Property
Public Function GetValue(Optional bInv As Variant) As Single
    If Not IsMissing(bInv) Then
        If bInv Then
            If m_SngVal = -1 Then GetValue = 2147483647
        Else
            If m_SngVal = -1 Then GetValue = 1
        End If
    Else
        GetValue = m_SngVal
    End If
End Function

Public Sub Move(PrW As Single, PrH As Single, brdr As Single, MinMax As Single, leftI As Long, nCnt As Long, bInv As Boolean)
    Dim L As Single, T As Single, W As Single, H As Single
    Dim S As Single
    'der Balken soll immer unten ausgerichtet sein (BaseBottom),
    'und gleichzeitig seine Größe in Bezug zu den anderen beibehalten
    If MinMax < 0 Then MinMax = 0
    If MinMax = 0 Then MinMax = 1
    If nCnt = 0 Then nCnt = 1
    'was wenn der kleinste Wert der beste ist, und das auch ausgedrückt werden soll?
    'also der kleinste Wert soll den Größten Balken haben
    'dann mit eins durch m_SngVal rechnen
    L = brdr + ((PrW - brdr) * leftI / nCnt)
    W = (PrW - ((nCnt + 1) * brdr)) / nCnt
    If bInv Then
        If m_SngVal = -1 Then
            S = 2147483647
            MinMax = S
        Else
            S = m_SngVal
        End If
        H = (PrH - 3 * brdr) / S * MinMax
    Else
        If m_SngVal = -1 Then
            S = 1
            MinMax = S
        Else
            S = m_SngVal
        End If
        H = (PrH - 3 * brdr) * S / MinMax
    End If
    'BaseTop:
    'T = 2 * brdr
    'BaseBottom:
    T = PrH - H - brdr
    If W > 0 And H > 0 Then m_PB.Move L, T, W, H
    m_PB.Cls
    m_PB.PSet (0, 0)
    m_PB.Print CStr(m_SngVal)
End Sub

