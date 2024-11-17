VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Balkendiagramm"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnNew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SArr(0 To 3, 0 To 3) As Single
Private Beams As Beams

Private Sub Form_Load()
  Dim i As Long
  Set Beams = New_Beams(Picture1)
  'For i = 0 To 3
  Call Beams.Add(New_Beam(Me, Picture1, "JS", i + 1, &HFF&)):    i = i + 1 'rot
  Call Beams.Add(New_Beam(Me, Picture1, "OM", i + 1, &HFFFF&)):  i = i + 1 'gelb
  Call Beams.Add(New_Beam(Me, Picture1, "SM", i + 1, &HFF00&)):  i = i + 1 'grün
  Call Beams.Add(New_Beam(Me, Picture1, "VB", i + 1, &HFF0000)): i = i + 1 'blau
  'Next
  Call RandomFillSngArray(SArr)
  Call FillValuesToBeams
End Sub

Private Sub BtnNew_Click()
  Call RandomFillSngArray(SArr())
  Call FillValuesToBeams
  Call Beams.Invalidate
End Sub

Private Sub FillValuesToBeams()
Dim i As Long
  For i = 1 To Beams.Count
    Beams.Beam(i).Value = SArr(0, i - 1)
  Next
End Sub
Private Sub RandomFillSngArray(SngArr() As Single)
Dim i As Long, j As Long
  Randomize
  For i = LBound(SngArr, 1) To UBound(SngArr, 1)
    For j = LBound(SngArr, 2) To UBound(SngArr, 2)
      SngArr(i, j) = 100 * Rnd
    Next
  Next
End Sub

Private Sub Form_Paint()
  Beams.Invalidate
End Sub

Private Sub Form_Resize()
Dim L As Single, T As Single, W As Single, H As Single
Dim brdr As Single
  brdr = 8 * 15
  L = Picture1.Left
  T = Picture1.Top
  W = Me.ScaleWidth - L - brdr
  H = Me.ScaleHeight - T - brdr
  If W > 0 And H > 0 Then Picture1.Move L, T, W, H
End Sub
