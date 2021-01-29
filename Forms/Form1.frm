VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "KDS8Reader"
   ClientHeight    =   13455
   ClientLeft      =   2325
   ClientTop       =   690
   ClientWidth     =   8640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   13455
   ScaleWidth      =   8640
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12615
      Left            =   120
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "Form1.frx":0CCA
      Top             =   840
      Width           =   8535
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   480
      Width           =   8175
   End
   Begin VB.OptionButton OptClipB 
      Caption         =   "ClipBoard"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton OptRows 
      Caption         =   "rowbased"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton OptCols 
      Caption         =   "vertical"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8160
      Picture         =   "Form1.frx":0CD0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "drag'drop *.syx files into the textbox"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_sr As KDS8SyxReader
Private m_IsHeadRead As Boolean
Private m_KDS8ProgList As MKDS8Types.ListOf_KORG_DS8_PROG
Private m_FNm As String
Private m_Err As String

Private Sub Trace(ByVal e As String)
    m_Err = m_Err & vbCrLf & e
End Sub
Private Sub Combo1_Click()
    UpdateView
End Sub
'Private Sub Command2_Click()
'    Dim aProg As KORG_DS8_PROG
'    aProg = MKDS8Reader.List_Get_Item(Combo1.List(Combo1.ListIndex))
'    Dim s As String
'    s = KORG_DS8_Prog_ToCBStr(aProg)
'    Clipboard.SetText s
'    Text1.Text = Clipboard.GetText
'    Clipboard.Clear
'    Clipboard.SetText s
'End Sub
'Private Sub Command1_Click()
'    m_FNm = App.Path & "\synthzone\B2.syx"
'    OpenSyXFile m_FNm
'End Sub
'Private Sub Command2_Click()
'    m_FNm = App.Path & "\synthzone\B3.syx"
'    OpenSyXFile m_FNm
'End Sub

Private Sub Form_Load()
    InitCharTable
    Text1.Text = ""
    Combo1.Text = ""
    Label1.Caption = ""
    OptRows.value = True
    '
    'FNm = "C:\Users\Oliver Meyer\Documents\KORG DS-8\MyNewSyX3.syx"
    'FNm = "C:\Users\Oliver Meyer\Documents\VB6\KORG_DS_8_SysEx\Original.syx"
    '2012_11_26 om: OK, was man jetzt tun sollte:
    'mal die original ds8-syx-Datei aus dem Netz laden und schauen ob man wenigstens grob die richtigen Daten hat.
    'die Namen der Patches sind schonmal unterschiedlich -> Vermutung es sind andere, oder ein german-brand?
    'Set m_sr = MNew.KDS8SyxReader(FNm)
    If Len(Command) > 0 Then
        Dim FNm As String: FNm = Command
        If Left(FNm, 1) = """" Then FNm = Mid$(FNm, 2)
        If Right(FNm, 1) = """" Then FNm = Left$(FNm, Len(FNm) - 1)
        OpenSyXFile FNm
    End If
End Sub

Private Sub ReadAll()
    If Not m_IsHeadRead Then
        If m_sr Is Nothing Then Exit Sub
        m_IsHeadRead = ReadHead(m_sr)
        If m_IsHeadRead Then
            m_sr.HeaderOffset = m_sr.FilePos - 1
            m_sr.Is8thByteSkip = True
        End If
    End If
    
try: On Error GoTo finally
    
    MKDS8Reader.List_Clear
    
    Dim aProg As KORG_DS8_PROG
    
    Dim i As Long
    For i = 0 To 99
        'Debug.Print CStr(i)
        aProg = MKDS8Reader.ReadKORGDS8_Prog(m_sr)
        Call MKDS8Reader.List_Add_KORG_DS8_PROG(aProg)
        'If i = 99 Then
        '    'Debug.Print aProg.VOICENAME.name
        'End If
    Next
    Call MKDS8Reader.List_ToComboBox(Combo1)
    Combo1.ListIndex = 0
    'Combo1.Text = Combo1.List(Combo1.ListIndex)
    'Loop
    Exit Sub
finally:
    'Close FNr
    m_sr.CClose
End Sub

Private Sub UpdateView()
    Dim aProg As KORG_DS8_PROG
    Dim i As Long: i = Combo1.ListIndex
    Label1.Caption = Format(i, "00")
    aProg = MKDS8Reader.List_Get_Item(Combo1.List(i))
    'Debug.Print MKDS8Reader.List_Count
    Text1.Text = KORG_DS8_Prog_ToStr(aProg, GetFormat)
End Sub
Function GetFormat() As EFmt
    Select Case True
    Case OptRows.value:  GetFormat = EFmt.rowbased
    Case OptCols.value:  GetFormat = EFmt.vertical
    Case OptClipB.value: GetFormat = EFmt.Clipboard
    Case Else: OptRows.value = True
    End Select
End Function
Private Function ReadHead(sr As KDS8SyxReader) As Boolean
    Dim b As Byte
try: On Error GoTo catch
    'SysEx?
    b = m_sr.ReadByte
    If b <> MIDI_SysEX Then
        Trace "Midi-SysEx format not found (Midi System Exclusive Messages) *.syx"
        GoTo catch
    End If
    'KORG-ID?
    b = m_sr.ReadByte
    If b <> ID_KORG Then
        Trace "Korg-ID not found"
        GoTo catch
    End If
    '30+x
    b = m_sr.ReadByte: If b < &H30 Then _
                            GoTo catch
    'DS8_ID 19 (&H13)
    b = m_sr.ReadByte
    If b <> ID_KORG_DS8 Then
        Trace "Kord DS-8-ID not found"
        GoTo catch
    End If
    '&H4C
    b = m_sr.ReadByte: If b <> &H4C Then _
                            GoTo catch
    '&H0
    b = m_sr.ReadByte: If b <> 0 Then _
                            GoTo catch
    ReadHead = True
    Exit Function
catch:
    ReadHead = False
    If Len(m_Err) > 0 Then MsgBox m_Err, vbInformation
    m_Err = ""
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Close #FNr
    Set m_sr = Nothing
End Sub

Private Sub OptClipB_Click()
    If Not m_sr Is Nothing Then UpdateView
End Sub
Private Sub OptCols_Click()
    If Not m_sr Is Nothing Then UpdateView
End Sub
Private Sub OptRows_Click()
    If Not m_sr Is Nothing Then UpdateView
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OLEDragDrop Data, Effect
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OLEDragDrop Data, Effect ', Button, Shift, X, Y
End Sub
Private Sub OLEDragDrop(Data As DataObject, Effect As Long) ', Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        m_FNm = Data.Files(1)
        OpenSyXFile m_FNm
    End If
End Sub
Private Sub OpenSyXFile(ByVal FNm As String)
    m_FNm = FNm
    m_IsHeadRead = False
    m_Err = vbNullString
    Set m_sr = MNew.KDS8SyxReader(m_FNm)
    ReadAll
    UpdateView
End Sub
Private Sub Form_Resize()
    Dim l As Single, T As Single, W As Single, H As Single
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    l = Me.ScaleWidth - Image1.Width
    Image1.Move l, 0
    l = 0: T = 4 * brdr: W = 3 * brdr: H = Combo1.Height
    If W > 0 And H > 0 Then Label1.Move l, T, W, H
    l = l + W: T = 4 * brdr
    W = Me.ScaleWidth - l: H = Combo1.Height
    If W > 0 And H > 0 Then Combo1.Move l, T, W ', H
    l = 0
    T = T + H + brdr
    W = Me.ScaleWidth - l
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move l, T, W, H
End Sub

