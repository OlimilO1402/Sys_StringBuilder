VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilderJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Buffer As String
Private BufferLen As Long
Private Pointer As Long

Friend Sub New_(Optional ByVal Value As String, Optional ByVal startIndex As Long, Optional ByVal Length As Long, Optional ByVal Capacity As Long, Optional ByVal maxCapacity As Long)
    'Set StringBuilderJS = New StringBuilderJS: StringBuilderJS.New_ Value, startIndex, Length, Capacity, maxCapacity
    Buffer = Value
End Sub

'Auf "Leerstring" setzen:
Public Sub Clear()
    Pointer = 1
End Sub

'String anhängen:
Public Sub Append(ByRef Value As String)
    Dim PointerNew As Long
    
    'Benötigten Buffer berechnen:
    PointerNew = Pointer + LenB(Value)
    'Ggf. Buffer vergrößern:
    If PointerNew > BufferLen Then
        Buffer = Buffer & Space$(PointerNew)
        BufferLen = LenB(Buffer)
    End If
    
    'String passend kopieren:
    MidB$(Buffer, Pointer) = Value
    Pointer = PointerNew
End Sub

Public Property Let Capacity(aCap As Long)
    'buffer =
End Property
Public Property Get Capacity() As Long
    Capacity = Len(Buffer)
End Property

'Länge des internen Strings:
Public Property Get Length() As Long
    Length = Pointer \ 2
End Property

'Inhalt des internen Strings:
Public Function ToStr() As String
    ToStr = LeftB$(Buffer, Pointer - 1)
End Function

'Neues Objekt initialisieren:
Private Sub Class_Initialize()
    Clear
    'Buffer = InternalAllocateStr(81920)
    'BufferLen = LenB(Buffer)
End Sub


