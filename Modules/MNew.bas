Attribute VB_Name = "MNew"
Option Explicit

'Jost Schwiders  StringBuilder www.VBTec.de 'originally it is called Concat
Public Function StringBuilderJS(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilderJS
    Set StringBuilderJS = New StringBuilderJS: StringBuilderJS.New_ Value, startIndex, Length, Capacity, maxCapacity
End Function

'Oliver Meyers  StringBuilder www.MBO-Ing.com
Public Function StringBuilder(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilder
    Set StringBuilder = New StringBuilder: StringBuilder.New_ Value, startIndex, Length, Capacity, maxCapacity
End Function

'Steve McMahons StringBuilder www.VBAccelerator.com
Public Function StringBuilderSM(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilderSM
    Set StringBuilderSM = New StringBuilderSM: StringBuilderSM.ChunkSize = Capacity
End Function

'only a class around VB's original String concatenation with the "&"-operator
Public Function StringBuilderVB(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long) As StringBuilderVB
    Set StringBuilderVB = New StringBuilderVB: StringBuilderVB.New_ Value, startIndex, Length, Capacity, maxCapacity
End Function

'##############################'   SBBench   '##############################'
Public Function SBBench(aFncPtr As Long, aSBBenchSituation As SBBenchSituation) As SBBench
    Set SBBench = New SBBench: SBBench.New_ aFncPtr, aSBBenchSituation
End Function

Public Function SBBenchSituation(Optional aXAvgRuns As Long = -1, Optional aStrVal As String = vbNullString, Optional aStrLen As Long = -1, Optional aNLoops As Long = -1) As SBBenchSituation
    Set SBBenchSituation = New SBBenchSituation: SBBenchSituation.New_ aXAvgRuns, aStrVal, aStrLen, aNLoops
End Function

'##############################'    Beams    '##############################'
Public Function Beams(aPB As PictureBox) As Beams
    Set Beams = New Beams: Beams.New_ aPB
End Function
Public Function Beam(aFrm As Form, aParent As PictureBox, aName As String, Index As Long, aColor As Long) As Beam
    Set Beam = New Beam: Beam.New_ aFrm, aParent, aName, Index, aColor
End Function
