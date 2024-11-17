Attribute VB_Name = "ModConstructors"
Option Explicit

Public Function New_Beam(aFrm As Form, aParent As PictureBox, aName As String, Index As Long, Optional aColor As Long = &HFF&) As Beam
  Set New_Beam = New Beam
  Call New_Beam.NewC(aFrm, aParent, aName, Index, aColor)
End Function

Public Function New_Beams(aPB As PictureBox) As Beams
  Set New_Beams = New Beams
  Call New_Beams.NewC(aPB)
End Function
