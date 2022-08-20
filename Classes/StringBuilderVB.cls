VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilderVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_StrVal As String

'This really simple Class is the native Visual Basic
'way to Append/Concat Strings
'rather slow!
Private Sub Class_Initialize()
    '
End Sub
Private Sub Class_Terminate()
    m_StrVal = vbNullString
End Sub

Friend Sub New_(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long)
    'Set StringBuilderVB = New StringBuilderVB: StringBuilderVB.New_ Value, startIndex, Length, Capacity, maxCapacity
    m_StrVal = Value
End Sub

Public Sub Append(aStrVal As String)
    m_StrVal = m_StrVal & aStrVal
End Sub

Public Property Get Length() As Long
    Length = Len(m_StrVal)
End Property

Public Function ToStr() As String
    ToStr = m_StrVal
End Function

