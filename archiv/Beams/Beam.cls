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
Private mPB As PictureBox
Private mSngVal As Single
Private mName As String
Private mColor As Long
Public Sub NewC(aFrm As Form, aParent As PictureBox, aName As String, Index As Long, Optional aColor As Long = &HFF&)
  Set mPB = aFrm.Controls.Add("VB.PictureBox", aName & CStr(Index), aParent)
  mName = aName
  mPB.Visible = True
  mPB.BackColor = aColor
  mPB.BorderStyle = 0
End Sub

Public Property Let Name(aName As String)
  mName = aName
End Property
Public Property Get Name() As String
  Name = mName
End Property

Public Property Let Color(aColor As String)
  mColor = aColor
End Property
Public Property Get Color() As String
  Color = mColor
End Property

'Public Property Get PB() As PictureBox
'  Set PB = mPB
'End Property

Public Property Let Value(aSngVal As Single)
  mSngVal = aSngVal
End Property
Public Property Get Value() As Single
  Value = mSngVal
End Property

Public Sub Move(PrW As Single, PrH As Single, brdr As Single, Max As Single, leftI As Long, nCnt As Long)
Dim L As Single, T As Single, W As Single, H As Single
  'der Balken soll immer unten ausgerichtet sein (BaseBottom),
  'und gleichzeitig seine Größe in Bezug zu den anderen beibehalten
  L = brdr + ((PrW - brdr) * leftI / nCnt)
  W = (PrW - ((nCnt + 1) * brdr)) / nCnt
  H = (PrH - 3 * brdr) * mSngVal / Max
  'BaseTop:
  'T = 2 * brdr
  'BaseBottom:
  T = PrH - H - brdr
  If W > 0 And H > 0 Then mPB.Move L, T, W, H
  mPB.Cls
  mPB.PSet (0, 0)
  mPB.Print CStr(mSngVal)
End Sub
