VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SBBench"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Diese Klasse stellt den Benchmark eines StringBuilder-Objektes dar
Implements IBench
'es treten vier Klassen gegeneinander an
'hier wird jeweils eine Klasse getestet
'jeder Test muß 10 mal laufen
'jede Klasse muß folgende Tests absolvieren:
'* Verknüpfung von Strings mit verschiedener Länge
'  1, 5, 20, 64
'* verschiedene Looplängen
'  100, 10000, 1000000
' je nach Größe des Strings maximale Stringgröße: 64Mb
'
Private mBenches() As Single 'enthält alle Zeitwerte
Private mFncPtr    As Long
Private mSit       As SBBenchSituation

Private Sub Class_Initialize()
  '
End Sub
Private Sub Class_Terminate()
    Erase mBenches
End Sub

Friend Sub New_(aFncPtr As Long, aBenchSituation As SBBenchSituation)
    mFncPtr = aFncPtr
    Set mSit = aBenchSituation
    ReDim mBenches(0 To mSit.XAvgRuns - 1, 0 To mSit.CountSLn - 1, 0 To mSit.CountLps - 1)
End Sub

Public Sub Run()
    Call IBench_Run
End Sub

Public Property Get OverallTime() As Single
    OverallTime = IBench_OverallTime
End Property
Private Property Get IBench_OverallTime() As Single
    Dim rv As Double, T As Variant
    For Each T In mBenches
        If T > 0 Then
            rv = rv + CDbl(T)
        End If
    Next
    IBench_OverallTime = CSng(rv)
End Property

Private Sub IBench_Run()
    Dim i As Long, j As Long, k As Long
    Dim n As Long, s As String
    Dim Mem As Single
    
    For k = 0 To mSit.CountLps - 1
        'unterschiedliche LoopLängen
        n = mSit.NLoops(k)
        For j = 0 To mSit.CountSLn - 1
            'unterschiedliche StringLängen
            If Len(mSit.StrVal) > 0 Then
                s = mSit.StrVal
            Else
                s = String$(mSit.StrLen(j), "x")
            End If
            Mem = MemoryUsedKiBy(n, s)
            If Mem < 64000 Then
                For i = 0 To mSit.XAvgRuns - 1 'um einen guten Mittelwert zu erhalten wird jeder Test 10 mal durchgeführt
                    'die Einzeldurchläufe, mit gleichen Parametern
                    mBenches(i, j, k) = BenchRunSB(mFncPtr, n, s)
                Next
            Else
                'ist der Wert < 0
                ' => der Test wurde nicht geführt
                For i = 0 To 9
                    mBenches(i, j, k) = -1
                Next
            End If
        Next
    Next
End Sub
Public Function GetTime(IxLoops As Long, IxStr As Long) As Single
    Dim i As Long
    For i = 0 To 9
        GetTime = GetTime + mBenches(i, IxStr, IxLoops)
    Next
    GetTime = GetTime \ 10
End Function
Private Function MemoryUsedKiBy(Loops As Long, aStrVal As String) As Single
    Dim rv As Double
    'gibt die Länge des kompletten Strings in KByte zurück
    rv = mSit.CountLps * (LenB(aStrVal) / 1024)
    MemoryUsedKiBy = CSng(rv)
    'StringLength = 2 * (2 ^ IndexLoopLen) * (10 ^ ((IndexStr + 1) * 2))
End Function

