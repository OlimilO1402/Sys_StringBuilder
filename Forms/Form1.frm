VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Benchmarking StringBuilder"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton OptShowCmplxBench 
      Caption         =   "ShowCmplxBench"
      Height          =   375
      Left            =   2040
      Style           =   1  'Grafisch
      TabIndex        =   23
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton OptShowSimpleTest 
      Caption         =   "ShowSimpleTest"
      Height          =   375
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox PnlSimpleTest 
      BorderStyle     =   0  'Kein
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5775
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   600
      Width           =   7815
      Begin VB.CommandButton BtnClearResults 
         Caption         =   "Clear"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.ListBox LstResults 
         Height          =   1035
         Left            =   4800
         TabIndex        =   15
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton BtnSBVB 
         Caption         =   "StringBuilderVB"
         Height          =   375
         Left            =   5760
         TabIndex        =   13
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton BtnSBSM 
         Caption         =   "StringBuilderSM"
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton BtnSBOM 
         Caption         =   "StringBuilderOM"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton BtnSBJS 
         Caption         =   "StringBuilderJS"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   3255
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   6
         Top             =   2280
         Width           =   7695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Text            =   "1000"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label LblN 
         Caption         =   "n := "
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "1, 100, 10000, 1000000, ..."
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Append/Concat    the following String:"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "How many loops:"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   840
         Width           =   3855
      End
   End
   Begin VB.PictureBox PnlCmplxBench 
      BorderStyle     =   0  'Kein
      Height          =   6015
      Left            =   120
      ScaleHeight     =   6015
      ScaleWidth      =   7695
      TabIndex        =   14
      Top             =   240
      Width           =   7695
      Begin VB.PictureBox PBBeams 
         BackColor       =   &H00FFFFFF&
         Height          =   3735
         Left            =   0
         ScaleHeight     =   3675
         ScaleWidth      =   6795
         TabIndex        =   21
         Top             =   600
         Width           =   6855
      End
      Begin VB.CommandButton BtnCSBVB 
         Caption         =   "ComplexB SBVB"
         Height          =   375
         Left            =   5760
         TabIndex        =   20
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton BtnCSBSM 
         Caption         =   "ComplexB SBSM"
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton BtnCSBOM 
         Caption         =   "ComplexB SBOM"
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton BtnCSBJS 
         Caption         =   "ComplexB SBJS"
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_AStr      As String
Private m_n         As Long
Private SW          As StopWatch
Private m_SBBenches As BenchCollection
Private m_Beams     As Beams
'Private Declare Sub Application_EnableVisualStyles Lib "comctl32.dll" Alias "InitCommonControls" ()

'im Startformular
Private Sub Form_Initialize()
    'Application_EnableVisualStyles
    OptShowSimpleTest.Value = True
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub BtnClearResults_Click()
    LstResults.Clear
End Sub

'Dies ist ein Testtreiber für ähnliche StringBuilder-Klassen, die je ein
'unterschiedliches Konzept für die Stringverknüpfung implementieren.
'Eines der leidigen Themen in Visual Basic, das oft die Gemüter erhitzt,
'und immer wieder für viel Diskussionsstoff in den Visual Basic Foren sorgt,
'war von jeher das Thema "die Verkettung von Strings".
'Jeder Visual Basic Programmierer hat gelernt, daß man, um Strings zu
'verketten, den "&"-Operator einsetzt. Doch nur derjenige weiß wie armselig
'langsam diese Art der Stringverkettung in Visual Basic abläuft, der schon
'einmal größere und längere Strings damit versucht hat zu verketten.
'Bei der Verkettung von kurzen Strings macht sich ein eventueller
'Geschwindigkeitsunterschied indes kaum bemerkbar.
'Doch bereits ab einer Gesamtstringlänge von 20 Zeichen, kann eine andere,
'durch einen anderen Algorithmus durchgeführte Stringverkettung enorme
'Geschwindigkeitsvorteile bieten.
'
'Um Stringverkettungen unter .NET durchzuführen gibt es dazu im Framework im
'Namespace mscorlib.System.Text die Klasse StringBuilder. Eine Stringverkettung
'ist allerdings auch mit dem "&"-Operator in VB.NET schon viel schneller als
'mit VB6.
'Was liegt also näher als die Idee aufzugreifen, die Stringverkettung durch
'eine Klasse durchführen zu lassen. Die Klasse könnte so gebaut werden, daß ein
'eventuelles Umrüsten eines VB6-Codes nach VB.NET keine Codekonvertierung mehr
'erforderlich macht. Das heißt also, daß die Klasse, wie die gleichnamige Klasse
'im .NET-Framework, die gleichen Prozeduren bereitstellt, und die gleiche
'Funktionalität bietet.
'Doch welche verschiedenen Verfahren bzw. Algorithmen zur StringVerkettung gibt
'es und welche sind in der tat schneller als der "&"-Operator?
'Da braucht man nicht lange zu googlen, schnell sind die Seiten von SteveMcMahon
'vbaccelerator.com und von Jost Schwider vbtec.de ausfinfdig gemacht, die beide
'ein komplett unterschiedliches Verfahren verwenden, und mit enormen Geschwindig-
'keitsvorteilen gegenüber der VB-Methode aufwarten.
'Doch worin unterscheiden sich die Klassen genau?
'Um dieses praxisrelevant beurteilen zu können, sind umfangreiche Tests
'erforderlich, und deshalb ist dies am besten in einer Benchmarksituation
'zu untersuchen.
'Auch kann man so die jeweiles besten Teile der Klassen aufspüren und in einer
'optimierten Klasse zusammenführen.
'Um eventuellen zukünftigen Fehlern bzw Inkompatibilitäten vorzubeugen, kann man
'die jeweils anderen Verfahren als Kommentare in die Klasse miteinbeziehen, sodaß
'in der Zukunft jederzeit zu einem anderen Verfahren gewechselt werden kann.
'
'##############################'   Form   '##############################'
Private Sub Form_Load()
    Dim SBBenchSituation As SBBenchSituation
    Set SW = New StopWatch
    'SW.Calibrate
    Text3.Text = "Dies ist nur irgend so ein ellenlanger String der hier steht. " & vbCrLf 'eine Länge von 62 Character + 2 Zeichen für vbCrLf
    Text2.Text = 10 '000
    'bei m_n = 100000
    'File: 5947kb
    
    Set SBBenchSituation = New SBBenchSituation
    
    Set m_SBBenches = New BenchCollection
    
    'Jost Schwiders 'originally Concat.Concat
    m_SBBenches.Add "JS", MNew.SBBench(AddressOf BenchRunSBJS, SBBenchSituation)
    
    'Oliver Meyers 'StringBuilder.Append
    m_SBBenches.Add "OM", MNew.SBBench(AddressOf BenchRunSBOM, SBBenchSituation)
    
    'Steve McMahons 'cStringBuilder.Append
    m_SBBenches.Add "SM", MNew.SBBench(AddressOf BenchRunSBSM, SBBenchSituation)
    
    'Visual Basics 'Wrapper around "&"-Operator
    m_SBBenches.Add "VB", MNew.SBBench(AddressOf BenchRunSBVB, SBBenchSituation)
    
    Set mBenches.TB = Text1
    
    '####################'   Diagram   '####################'
    Set m_Beams = MNew.Beams(PBBeams)
    m_Beams.Inverted = True
    Dim i As Long
    m_Beams.Add MNew.Beam(Me, PBBeams, "JS", i + 1, &HFF&):   i = i + 1 'rot
    m_Beams.Add MNew.Beam(Me, PBBeams, "OM", i + 1, &HFFFF&): i = i + 1 'gelb
    m_Beams.Add MNew.Beam(Me, PBBeams, "SM", i + 1, &HFF00&): i = i + 1 'grün
    m_Beams.Add MNew.Beam(Me, PBBeams, "VB", i + 1, &HFF0000): i = i + 1 'blau

End Sub

Private Sub Form_Paint()
    m_Beams.Invalidate
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    'Me.ScaleMode = vbPixels
    L = PnlSimpleTest.Left: T = PnlSimpleTest.Top
    W = Me.ScaleWidth - L - brdr
    H = Me.ScaleHeight - T - brdr
    If W > 0 And H > 0 Then PnlSimpleTest.Move L, T, W, H
    
    'L = PnlCmplxBench.Left: T = PnlCmplxBench.Top
    W = Me.ScaleWidth - L - brdr
    H = Me.ScaleHeight - T - brdr
    If W > 0 And H > 0 Then PnlCmplxBench.Move L, T, W, H
    
    L = Text1.Left: T = Text1.Top
    W = PnlSimpleTest.ScaleWidth '- L - Brdr
    H = PnlSimpleTest.ScaleHeight - T '- Brdr
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
      
End Sub

Private Sub PnlCmplxBench_Resize()
    Dim L As Single: L = PBBeams.Left
    Dim T As Single: T = PBBeams.Top
    Dim W As Single: W = PnlCmplxBench.ScaleWidth
    Dim H As Single: H = PnlCmplxBench.ScaleHeight
    If W > 0 And H > 0 Then PBBeams.Move L, T, W, H
End Sub

Private Sub OptShowCmplxBench_Click()
    PnlCmplxBench.ZOrder 0
    Set mBenches.TB = Nothing
    m_Beams.Invalidate
End Sub
Private Sub OptShowSimpleTest_Click()
    PnlSimpleTest.ZOrder 0
    Set mBenches.TB = Text1
End Sub
'##############################'   ComplexBenchButtons   '##############################'
Private Sub BtnCSBJS_Click()
    m_SBBenches.BenchMark("JS").Run
    Dim time As Single: time = m_SBBenches.BenchMark("JS").OverallTime
    MsgBox CStr(time)
    m_Beams.Beam(1).Value = time
    m_Beams.Invalidate
End Sub
Private Sub BtnCSBOM_Click()
    m_SBBenches.BenchMark("OM").Run
    Dim time As Single: time = m_SBBenches.BenchMark("OM").OverallTime
    MsgBox CStr(time)
    m_Beams.Beam(2).Value = time
    m_Beams.Invalidate
End Sub
Private Sub BtnCSBSM_Click()
    m_SBBenches.BenchMark("SM").Run
    Dim time As Single: time = m_SBBenches.BenchMark("SM").OverallTime
    MsgBox CStr(time)
    m_Beams.Beam(3).Value = time
    m_Beams.Invalidate
End Sub
Private Sub BtnCSBVB_Click()
    m_SBBenches.BenchMark("VB").Run
    Dim time As Single: time = m_SBBenches.BenchMark("VB").OverallTime
    MsgBox CStr(time)
    m_Beams.Beam(1).Value = time
    m_Beams.Invalidate
End Sub

Private Sub Text2_Change()
    If IsNumeric(Text2.Text) Then m_n = CLng(Text2.Text)
    CalcMemoryUsed
End Sub

Private Sub Text3_Change()
    m_AStr = Text3.Text
    Label4.Caption = "Len: " & Len(m_AStr)
    CalcMemoryUsed
End Sub

Private Sub CalcMemoryUsed()
    Label3.Caption = "String uses memory: " & Format((m_n * LenB(m_AStr) / 1024), "###,###,###.0000") & "kb"
End Sub

'##############################'   The Simple Tests   '##############################'
Private Sub BtnSBJS_Click()
    Dim aBench As SBBench: Set aBench = MNew.SBBench(AddressOf BenchRunSBJS, MNew.SBBenchSituation(1, m_AStr, , m_n))
    aBench.Run
    LstResults.AddItem "JS: " & CStr(aBench.OverallTime) & "milliseconds"
End Sub

Private Sub BtnSBOM_Click()
    Dim aBench As SBBench: Set aBench = MNew.SBBench(AddressOf BenchRunSBOM, MNew.SBBenchSituation(1, m_AStr, , m_n))
    aBench.Run
    LstResults.AddItem "OM: " & CStr(aBench.OverallTime) & "milliseconds"
End Sub

Private Sub BtnSBSM_Click()
    Dim aBench As SBBench: Set aBench = MNew.SBBench(AddressOf BenchRunSBSM, MNew.SBBenchSituation(1, m_AStr, , m_n))
    aBench.Run
    LstResults.AddItem "SM: " & CStr(aBench.OverallTime) & "milliseconds"
End Sub

Private Sub BtnSBVB_Click()
    Dim aBench As SBBench
    If MsgBox("Oh really, this may take a while, are you sure?", vbOKCancel) = vbOK Then
        Set aBench = MNew.SBBench(AddressOf BenchRunSBVB, MNew.SBBenchSituation(1, m_AStr, , m_n))
        'Set aBench = New_SBBench(AddressOf BenchRunSBnet, New_SBBenchSituation(1, m_AStr, , m_n))
        aBench.Run
        LstResults.AddItem "VB: " & CStr(aBench.OverallTime) & "milliseconds"
    End If
End Sub

Private Sub BtnSBNET_Click()
    Dim aBench As SBBench: Set aBench = MNew.SBBench(AddressOf BenchRunSBVB, MNew.SBBenchSituation(1, m_AStr, , m_n))
    aBench.Run
    LstResults.AddItem "VB: " & CStr(aBench.OverallTime) & "milliseconds"
End Sub

Private Sub WriteToFile(PFN As String, AStr As String)
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Write As FNr
    Put FNr, , AStr
    Close FNr
End Sub

