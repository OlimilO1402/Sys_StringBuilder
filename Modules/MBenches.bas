Attribute VB_Name = "MBenches"
Option Explicit
Public Declare Function BenchRunSB Lib "user32.dll" Alias "CallWindowProcA" (ByVal FncPtr As Long, ByVal LoopLen As Long, ByVal StrVal As String, Optional ByVal ValA As Long, Optional ByVal ValB As Long) As Single
Public SW As New StopWatch
Public TB As TextBox

Public Function BenchRunSBJS(ByVal LoopLen As Long, ByVal StrVal As String, Optional ByVal ValA As Long, Optional ByVal ValB As Long) As Single
    Dim aStrVal As String: aStrVal = StrConv(StrVal, vbUnicode)
    Dim aSB As StringBuilderJS: Set aSB = MNew.StringBuilderJS(vbNullString, , , 16777216)
    SW.Reset
    SW.Start
    Dim i As Long
    For i = 0 To LoopLen - 1
        aSB.Append (aStrVal)
    Next
    SW.SStop
    If Not TB Is Nothing Then TB.Text = aSB.ToStr
    BenchRunSBJS = CSng(SW.ElapsedMilliseconds)
End Function

Public Function BenchRunSBOM(ByVal LoopLen As Long, ByVal StrVal As String, Optional ByVal ValA As Long, Optional ByVal ValB As Long) As Single
    Dim aStrVal As String: aStrVal = StrConv(StrVal, vbUnicode)
    Dim aSB As StringBuilder: Set aSB = MNew.StringBuilder(vbNullString, , , 16777216)
    SW.Reset
    SW.Start
    Dim i As Long
    For i = 0 To LoopLen - 1
      Call aSB.Append(aStrVal)
    Next
    SW.SStop
    If Not TB Is Nothing Then TB.Text = aSB.ToStr
    BenchRunSBOM = CSng(SW.ElapsedMilliseconds)
End Function

Public Function BenchRunSBSM(ByVal LoopLen As Long, ByVal StrVal As String, Optional ByVal ValA As Long, Optional ByVal ValB As Long) As Single
    Dim aStrVal As String: aStrVal = StrConv(StrVal, vbUnicode)
    Dim aSB As StringBuilderSM: Set aSB = MNew.StringBuilderSM(vbNullString, , , 16777216)
    SW.Reset
    SW.Start
    Dim i As Long
    For i = 0 To LoopLen - 1
      Call aSB.Append(aStrVal)
    Next
    SW.SStop
    If Not TB Is Nothing Then TB.Text = aSB.ToStr
    BenchRunSBSM = CSng(SW.ElapsedMilliseconds)
End Function

Public Function BenchRunSBVB(ByVal LoopLen As Long, ByVal StrVal As String, Optional ByVal ValA As Long, Optional ByVal ValB As Long) As Single
    Dim aStrVal As String: aStrVal = StrConv(StrVal, vbUnicode)
    Dim aSB As StringBuilderVB: Set aSB = MNew.StringBuilderVB(vbNullString, , , 16777216)
    SW.Reset
    SW.Start
    Dim i As Long
    For i = 0 To LoopLen - 1
        Call aSB.Append(aStrVal)
    Next
    SW.SStop
    If Not TB Is Nothing Then TB.Text = aSB.ToStr
    BenchRunSBVB = CSng(SW.ElapsedMilliseconds)
End Function

'Hmmm, ich weiß nicht geht das mit VBFusion ??? - Nö
'Public Function BenchRunSBnet(ByVal LoopLen As Long, ByVal StrVal As String, Optional ByVal ValA As Long, Optional ByVal ValB As Long) As Single
'Dim i As Long
'Dim aStrVal As String: aStrVal = StrConv(StrVal, vbUnicode)
'Dim aSB As New mscorlib.StringBuilder ': Set aSB = New_StringBuilderVB(vbNullString, , , 16777216)
'  SW.Calibrate
'  SW.Start
'  For i = 0 To LoopLen - 1
'    Call aSB.Append(aStrVal)
'  Next
'  SW.Halt
'  If Not TB Is Nothing Then TB.Text = aSB.ToStr
'  BenchRunSBnet = CSng(SW.TimeSpan)
'  Set aSB = Nothing
'End Function
