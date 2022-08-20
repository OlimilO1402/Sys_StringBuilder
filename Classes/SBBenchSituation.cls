VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SBBenchSituation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse stellt eine Benchmarksituation dar.
'Dies ist also eine sehr spezialisierte Klasse
'die hier nur für den StringBuilderBenchmark ist.
'Diese Klasse definiert also einen Benchmark, den
'jedes StringBuilder-Objekt abarbeiten muß.
'diese Klasse wird am besten als Singletonobj benutzt,
'd.h. nur ein Objekt dieser Klasse muß vorhanden sein.
'
'jeder Test soll X-mal laufen, um einen Mittelwert bestimmen zu können.
Public XAvgRuns   As Long
'ein facher Testdurchlauf (aus 10)
'besteht aus einem Loop durch verschiedene StringLängen die
'verschieden oft aneinandergereiht werden sollen
Public CountSLn   As Long
Private mStrLen() As Long 'die verschiedenen Stringlängen
Public CountLps   As Long
Private mNLoops() As Long 'die Anzahl an Aneinanderreihungen
Public StrVal     As String 'dieser String wird aneinandergereiht

Private Sub Class_Initialize()
    XAvgRuns = 10
    CountSLn = 4
    CountLps = 4
    ReDim mStrLen(0 To CountSLn - 1)
    ReDim mNLoops(0 To CountLps - 1)
'
'* Looplänge 'wie oft der String angehängt wird
'   100, 10000, 1000000, 100000000;
'  10^2,  10^4,    10^6,      10^8;
'    10,  1000,  100000,  1000000
'
'* Stringlänge
'    1,   4,  16,  64;
'  2^0, 2^2, 2^4, 2^8;
'  maximale getestete StringLänge: 64Mb
'
'Dim i As Long

'  For i = 0 To 3
'    'unterschiedliche StringLängen
'    If i = 0 Then
'      mStrLen(i) = 1
'    Else
'      mStrLen(i) = 2 ^ (2 ^ i)
'    End If
'    'mLoops(i) = 10 ^ (i + i + 2)
'  Next
    mStrLen(0) = 1
    mStrLen(1) = 16
    mStrLen(2) = 32
    mStrLen(3) = 64
    
    mNLoops(0) = 10 '0
    mNLoops(1) = 10 '00
    mNLoops(2) = 10 '000
    mNLoops(3) = 10 '0000

End Sub

Friend Sub New_(Optional aXAvgRuns As Long = -1, Optional aStrVal As String = vbNullString, Optional aStrLen As Long = -1, Optional aNLoops As Long = -1)
    If XAvgRuns > 0 Then XAvgRuns = aXAvgRuns
    If Len(aStrVal) > 0 Then StrVal = aStrVal
    CountSLn = 1
    ReDim mStrLen(CountSLn - 1)
    If aStrLen > 0 Then
        mStrLen(CountSLn - 1) = aStrLen
    End If
    CountLps = 1
    ReDim mNLoops(CountLps - 1)
    If aNLoops > 0 Then
        mNLoops(CountLps - 1) = aNLoops
    End If
End Sub

Public Property Get StrLen(Index As Long) As Long
    StrLen = mStrLen(Index)
End Property

Public Property Get NLoops(Index As Long) As Long
    NLoops = mNLoops(Index)
End Property

