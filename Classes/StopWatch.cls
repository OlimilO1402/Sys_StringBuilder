VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class StopWatch
Private Type LongInt
    Lo As Long
    Hi As Long
End Type
'Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Double) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LongInt) As Long
'Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Double) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LongInt) As Long
Private mCalib As Double
Private mStrt  As LongInt
Private mFini  As LongInt
Private mFreq  As LongInt

Private Sub Class_Initialize()
    QueryPerformanceFrequency mFreq
    Calibrate
End Sub

Public Sub Calibrate()
    QueryPerformanceCounter mStrt
    QueryPerformanceCounter mFini
    mCalib = (D(mFini) - D(mStrt)) / D(mFreq) * 1000
End Sub

Public Sub Halt()
    QueryPerformanceCounter mFini
End Sub

Public Sub Start()
    QueryPerformanceCounter mStrt
End Sub

Public Property Get TimeSpan() As Double 'in milliseconds
    TimeSpan = (D(mFini) - D(mStrt)) / D(mFreq) * 1000 - mCalib
End Property

Public Function ToStr() As String
    ToStr = Format(TimeSpan, "#.000,000") & "milliseconds"
End Function

Private Function D(LX As LongInt) As Double
    Dim DH As Double:    DH = LX.Hi
    Dim DL As Double:    DL = LX.Lo
    If DL < 0 Then DL = 4294967296# + DL + 1
    If DH < 0 Then DH = 4294967296# + DH + 1
    D = DL + DH * 4294967296#
End Function

