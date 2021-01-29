Attribute VB_Name = "MKDS8StrWriter"
Option Explicit
'umschaltbar zwischen vertikaler und zeilenförmiger Darstellung?
Enum EFmt
    vertical
    rowbased
    Clipboard
End Enum
Dim s1 As String
Dim s2 As String
Dim m_fmt As EFmt
Dim CharTable() As String

Sub InitCharTable()
    Dim i As Integer
    ReDim ct(0 To 95) As String
    ct(i) = " ": i = i + 1: ct(i) = "!": i = i + 1: ct(i) = """": i = i + 1: ct(i) = "#": i = i + 1: ct(i) = "$": i = i + 1: ct(i) = "%": i = i + 1: ct(i) = "&": i = i + 1: ct(i) = "'": i = i + 1
    ct(i) = "(": i = i + 1: ct(i) = ")": i = i + 1: ct(i) = "*": i = i + 1: ct(i) = "+": i = i + 1: ct(i) = ",": i = i + 1: ct(i) = "-": i = i + 1: ct(i) = ".": i = i + 1: ct(i) = "/": i = i + 1
    ct(i) = "0": i = i + 1: ct(i) = "1": i = i + 1: ct(i) = "2": i = i + 1: ct(i) = "3": i = i + 1: ct(i) = "4": i = i + 1: ct(i) = "5": i = i + 1: ct(i) = "6": i = i + 1: ct(i) = "7": i = i + 1
    ct(i) = "8": i = i + 1: ct(i) = "9": i = i + 1: ct(i) = ":": i = i + 1: ct(i) = ";": i = i + 1: ct(i) = "<": i = i + 1: ct(i) = "=": i = i + 1: ct(i) = ">": i = i + 1: ct(i) = "?": i = i + 1
    ct(i) = "@": i = i + 1: ct(i) = "A": i = i + 1: ct(i) = "B": i = i + 1: ct(i) = "C": i = i + 1: ct(i) = "D": i = i + 1: ct(i) = "E": i = i + 1: ct(i) = "F": i = i + 1: ct(i) = "G": i = i + 1
    ct(i) = "H": i = i + 1: ct(i) = "I": i = i + 1: ct(i) = "J": i = i + 1: ct(i) = "K": i = i + 1: ct(i) = "L": i = i + 1: ct(i) = "M": i = i + 1: ct(i) = "N": i = i + 1: ct(i) = "O": i = i + 1
    ct(i) = "P": i = i + 1: ct(i) = "Q": i = i + 1: ct(i) = "R": i = i + 1: ct(i) = "S": i = i + 1: ct(i) = "T": i = i + 1: ct(i) = "U": i = i + 1: ct(i) = "V": i = i + 1: ct(i) = "W": i = i + 1
    ct(i) = "X": i = i + 1: ct(i) = "Y": i = i + 1: ct(i) = "Z": i = i + 1: ct(i) = "[": i = i + 1: ct(i) = "Y": i = i + 1: ct(i) = "]": i = i + 1: ct(i) = "^": i = i + 1: ct(i) = "_": i = i + 1
    ct(i) = "`": i = i + 1: ct(i) = "a": i = i + 1: ct(i) = "b": i = i + 1: ct(i) = "c": i = i + 1: ct(i) = "d": i = i + 1: ct(i) = "e": i = i + 1: ct(i) = "f": i = i + 1: ct(i) = "g": i = i + 1
    ct(i) = "h": i = i + 1: ct(i) = "i": i = i + 1: ct(i) = "j": i = i + 1: ct(i) = "k": i = i + 1: ct(i) = "l": i = i + 1: ct(i) = "m": i = i + 1: ct(i) = "n": i = i + 1: ct(i) = "o": i = i + 1
    ct(i) = "p": i = i + 1: ct(i) = "q": i = i + 1: ct(i) = "r": i = i + 1: ct(i) = "s": i = i + 1: ct(i) = "t": i = i + 1: ct(i) = "u": i = i + 1: ct(i) = "v": i = i + 1: ct(i) = "w": i = i + 1
    ct(i) = "x": i = i + 1: ct(i) = "y": i = i + 1: ct(i) = "z": i = i + 1: ct(i) = "{": i = i + 1: ct(i) = "|": i = i + 1: ct(i) = "}": i = i + 1: ct(i) = "~": i = i + 1: ct(i) = "": i = i + 1
    'matches nearly everywhere with ascii
    'except:
    'ct(60) = not "\" it is Yen-Zeichen
    'ct(94) = not "~" it is <= 'Pfeil links
    'ct(95) = not "" it is => 'Pfeil rechts
    
    CharTable = ct
End Sub
Function CompareCharTable() As Boolean
    Dim i As Integer
    Dim c As String
    For i = 0 To UBound(CharTable)
        c = Chr(i + 32)
        If CharTable(i) = c Then
            'Debug.Print "OK: " & CStr(i) & " : " & CharTable(i) & " = " & c
            CompareCharTable = True
        Else
            'Debug.Print "Oh no:" & CStr(i) & " : " & CharTable(i) & " <> " & c
            CompareCharTable = False
            'Exit Function
        End If
    Next
End Function
Function PadLeft(value As String, ByVal totallength As Long, Optional ByVal padchar As String = " ") As String
    'füllt links mit padchar auf
    Dim d As Long: d = totallength - Len(value)
    If d > 0 Then
        If padchar = " " Then
            PadLeft = Space(totallength)
            RSet PadLeft = value
        Else
            PadLeft = String(d, padchar) & value
        End If
    Else
        PadLeft = value
    End If
End Function
Function PadRight(value As String, ByVal totallength As Long, Optional ByVal padchar As String = " ") As String
    'füllt rechts mit padchar auf
    Dim d As Long: d = totallength - Len(value)
    If d > 0 Then
        If padchar = " " Then
            PadRight = Space(totallength)
            LSet PadRight = value
        Else
            PadRight = value & String(d, padchar)
        End If
    Else
        PadRight = value
    End If
End Function
Public Function Max(v1, v2)
    If v1 > v2 Then Max = v1 Else Max = v2
End Function
Public Function KORG_DS8_Prog_ToStr(aKORGDS8Prog As KORG_DS8_PROG, ByVal fmt As EFmt) As String
    Dim s As String: s = "KORG DS8 MIDI-SysEx" & vbCrLf & "==================="
    Dim i As Long: i = 1
    m_fmt = fmt
    With aKORGDS8Prog
        Select Case m_fmt
        Case vertical
            
            s = s & vbCrLf & "PITCH" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & PITCH_ToStr(.PITCH) & vbCrLf              'As PITCH        '3
            s = s & vbCrLf & "PITCH EG" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & PITCH_EG_ToStr(.PITCH_EG) & vbCrLf        'As PICTH_EG     '6  '9
            
            s = s & vbCrLf & "WFRM1" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & WFRM1_ToStr(.OSC1_WFRM1) & vbCrLf           'As WFRM         '5  '14
            s = s & vbCrLf & "WFRM2" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & WFRM2_ToStr(.OSC2_WFRM2) & vbCrLf           'As WFRM         '5  '19
            
            With .TIMBRE_EG
                s = s & vbCrLf & "T.EG1" & " " & ParamSetNr_ToStr(i): i = i + 1
                s = s & T_EG_ToStr(.OSC1) & vbCrLf      'As T_EG         '7  '26
                s = s & vbCrLf & "T.EG2" & " " & ParamSetNr_ToStr(i): i = i + 1
                s = s & T_EG_ToStr(.OSC2) & vbCrLf      'As T_EG         '7  '33
            End With
            
            With .AMPLIT_EG
                s = s & vbCrLf & "A.EG1" & " " & ParamSetNr_ToStr(i): i = i + 1
                s = s & A_EG_ToStr(.OSC1) & vbCrLf      'As A_EG         '6  '39
                s = s & vbCrLf & "A.EG2" & " " & ParamSetNr_ToStr(i): i = i + 1
                s = s & A_EG_ToStr(.OSC2) & vbCrLf      'As A_EG         '6  '45
            End With
            
            s = s & vbCrLf & "MG" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & MG_ToStr(.MODULATION_GEN) & vbCrLf        'As MG           '7  '52
            s = s & vbCrLf & "PORTAMENTO" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & PORTAMENTO_ToStr(.PORTAMENTO) & vbCrLf    'As PORTAMENTO   '2  '54
            s = s & vbCrLf & "JOYSTICK" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & JOYSTICK_ToStr(.JOYSTICK) & vbCrLf        'As JOYSTICK     '3  '57
            
            s = s & vbCrLf & "VELOCITY" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & VELOCITY_ToStr(.VELOCITY) & vbCrLf        'As VELOCITY     '4  '61
            s = s & vbCrLf & "AFT TOUCH" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & AFT_TOUCH_ToStr(.AFTER_TOUCH) & vbCrLf    'As AFT_TOUCH    '4  '65
            
            s = s & vbCrLf & "ASSIGNMODE" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & ASSIGNMODE_ToStr(.ASSIGN_MODE) & vbCrLf   'As ASSIGNMODE   '1  '66
            s = s & vbCrLf & "VOICENAME" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & VOICENAME_ToStr(.VOICENAME) & vbCrLf      'As VOICENAME    '10 '76
            
            s = s & vbCrLf & "MULTI EFFECT" & " " & ParamSetNr_ToStr(i): i = i + 1
            s = s & MULTI_EFFECT_ToStr(.MULTIEFFECT) & vbCrLf 'As MULTI_EFFECT '6  '82
                                                          '  es dürfen aber nur 80 sein
        Case rowbased
            Dim hl As Long: hl = 13
            
            s1 = vbCrLf & PadRight("PITCH", hl):        s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call PITCH_ToStr(.PITCH):                   s = s & s1 & vbCrLf & s2 & vbCrLf
            s1 = vbCrLf & PadRight("PITCH EG", hl):     s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call PITCH_EG_ToStr(.PITCH_EG):             s = s & s1 & vbCrLf & s2 & vbCrLf
            
            s1 = vbCrLf & PadRight("WFRM1", hl):        s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call WFRM1_ToStr(.OSC1_WFRM1):              s = s & s1 & vbCrLf & s2 & vbCrLf
            s1 = vbCrLf & PadRight("WFRM2", hl):        s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call WFRM2_ToStr(.OSC2_WFRM2):              s = s & s1 & vbCrLf & s2 & vbCrLf
            
            With .TIMBRE_EG
                s1 = vbCrLf & PadRight("T.EG1", hl):        s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
                Call T_EG_ToStr(.OSC1):                     s = s & s1 & vbCrLf & s2 & vbCrLf
                s1 = vbCrLf & PadRight("T.EG2", hl):        s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
                Call T_EG_ToStr(.OSC2):                     s = s & s1 & vbCrLf & s2 & vbCrLf
            End With
            
            With .AMPLIT_EG
                s1 = vbCrLf & PadRight("A.EG1", hl):        s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
                s = s & A_EG_ToStr(.OSC1):        s = s & s1 & vbCrLf & s2 & vbCrLf
                s1 = vbCrLf & PadRight("A.EG2", hl):        s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
                s = s & A_EG_ToStr(.OSC2):        s = s & s1 & vbCrLf & s2 & vbCrLf
            End With
            
            s1 = vbCrLf & PadRight("MG", hl):           s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call MG_ToStr(.MODULATION_GEN):             s = s & s1 & vbCrLf & s2 & vbCrLf
            s1 = vbCrLf & PadRight("PORTAMENTO", hl):   s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call PORTAMENTO_ToStr(.PORTAMENTO):         s = s & s1 & vbCrLf & s2 & vbCrLf
            s1 = vbCrLf & PadRight("JOYSTICK", hl):     s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call JOYSTICK_ToStr(.JOYSTICK):             s = s & s1 & vbCrLf & s2 & vbCrLf
            
            s1 = vbCrLf & PadRight("VELOCITY", hl):     s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call VELOCITY_ToStr(.VELOCITY):             s = s & s1 & vbCrLf & s2 & vbCrLf
            s1 = vbCrLf & PadRight("AFT TOUCH", hl):    s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call AFT_TOUCH_ToStr(.AFTER_TOUCH):         s = s & s1 & vbCrLf & s2 & vbCrLf
            
            s1 = vbCrLf & PadRight("ASSIGNMODE", hl):   s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call ASSIGNMODE_ToStr(.ASSIGN_MODE):        s = s & s1 & vbCrLf & s2 & vbCrLf
            s1 = vbCrLf & PadRight("VOICENAME", hl):    s2 = PadRight(ParamSetNr_ToStr(i), hl): i = i + 1
            Call VOICENAME_ToStr(.VOICENAME):           s = s & s1 & vbCrLf & s2 & vbCrLf
            s1 = vbCrLf & PadRight("MULTI EFFECT", hl): s2 = PadRight(ParamSetNr_ToStr(i) & " " & KDS8_EffectType_ToStr(.MULTIEFFECT.EffectType) & " ", hl): i = i + 1
            Call MULTI_EFFECT_ToStr(.MULTIEFFECT):      s = s & s1 & vbCrLf & s2 & vbCrLf
            
        End Select
    End With
    KORG_DS8_Prog_ToStr = s
End Function
Function ParamSetNr_ToStr(ByVal n As Long) As String
    Dim s As String
    Dim i As Long: i = n
    If i < 10 Then
        s = " " & CStr(i)
    Else
        i = n - 9
        s = "0" & CStr(i)
    End If
    ParamSetNr_ToStr = s
End Function
Function Param_ToStr(ByVal name As String, ByVal value As String) As String
    Select Case m_fmt
    Case EFmt.vertical
        Param_ToStr = vbCrLf & "    " & name & ": " & value
    Case EFmt.rowbased
        Dim l As Long: l = Max(Len(name), Len(value))
        s1 = s1 & PadRight(name, l + 1)
        s2 = s2 & PadRight(value, l + 1)
    End Select
End Function

Function OSC_ToStr(ByVal b As Byte) As String
    OSC_ToStr = IIf(b = 0, "0.5", Nibble_ToStr(b))
'    Dim s As String: s = IIf(b = 0, "0.5", Nibble_ToStr(b))
'    s = Nibble_ToStr(b)
'    If b = 0 Then s = "0.5"
'
'    Select Case b
'    Case 0:    s = "0.5"
'    Case Else: s = Nibble_ToStr(b)
'    End Select
'    OSC_ToStr = s
End Function
Function PITCH_ToStr(aPITCH As PITCH) As String
    Dim s As String
    With aPITCH
        s = s & Param_ToStr("OSC1", OSC_ToStr(.OSC1))   'As Byte 'Oscillator 1 ' 0.5; . . . ; 15;   0.5 (16-Fuss-Orgel); 1 (8-FO); 2 (4-FO); . . .
        s = s & Param_ToStr("OSC2", OSC_ToStr(.OSC2))   'As Byte 'Oscillator 1 ' 0.5; . . . ; 15;   0.5 (16-Fuss-Orgel); 1 (8-FO); 2 (4-FO); . . .
        s = s & Param_ToStr("DTN", SemiNib_ToStr(.DTN)) 'As Byte 'Detune       ' 0; 1; 2; 3
    End With
    PITCH_ToStr = s
End Function
Function PITCH_EG_ToStr(aPITCH_EG As PITCH_EG) As String
    Dim s As String
    With aPITCH_EG
        s = s & Param_ToStr("STL ", Signed_7Bit_ToStr(.STL))   'Anfangspegel         '-63; . . . ; +63; 'Die Tonhöhe bei der die Klangfarbe beginnt jedesmal wenn eine Note gespielt wird
        s = s & Param_ToStr("ATK ", Unsigned_6Bit_ToStr(.ATK)) 'Einschwingungszeit   '  0; . . . ;  63; 'Die Zeit in der sich die Tonhöhe vom Anfangspegel zum einschwingungspegel ändert
        s = s & Param_ToStr("ATL ", Signed_7Bit_ToStr(.ATL))   'Einschwingungspegel  '-63; . . . ; +63; 'Der Spitzenpegelwert der Tonhöhe
        s = s & Param_ToStr("DEC ", Unsigned_6Bit_ToStr(.DEC)) 'Ausschwingungszeit   '  0; . . . ;  63; 'Die Zeit in der die Tonhöhe vom Einschwingungspegel zum Normalpegel zurückkehrt, währen die Taste gedrücktgehalten bleibt.
        s = s & Param_ToStr("REL ", Unsigned_6Bit_ToStr(.REL)) 'Abklingung           '  0; . . . ;  63; 'Die Zeit in der die Tonhöht zum Abklingungspegel wechselt, nachdem die Taste losgelassen wurde.
        s = s & Param_ToStr("RLL ", Signed_7Bit_ToStr(.RLL))   'Abklingungspegel     '-63; . . . ; +63; 'Der Pegel zu dem die Tonhöhe wechselt, nachdem die Taste losgelassen wurde.
    End With
    PITCH_EG_ToStr = s
End Function

Public Function KDS8_ONOFF_ToStr(ByVal e As KDS8_ONOFF) As String
    Dim s As String
    Select Case e
    Case kds8_OFF: s = "OFF"
    Case kds8_ON:  s = "ON"
    Case Else: s = "&H" & Hex(e)
    End Select
    KDS8_ONOFF_ToStr = s
End Function

Function WFRM1_ToStr(aWFRM1 As WFRM1) As String
    Dim s As String
    With aWFRM1
                                                                                                 '             ' SAW,         RECT,        LiSAW,              LiRECT
        s = s & Param_ToStr("TYP  ", SemiNib1_ToStr(.TYP))        'As Byte 'Wellenformart        '1; 2; 3; 4; 1: Sägezahn; 2: Rechteck; 3: Heller Sägezahn; 4: Helles Rechteck;
        s = s & Param_ToStr("SPCT ", Unsigned_3Bit1_ToStr(.SPCT)) 'As Byte 'Spektrum             '1; . . .; 8;
        s = s & Param_ToStr("RING ", SemiNib_ToStr(.RING))        'As Byte 'Glocken - Modulation '0; 1; 2; 3;
        s = s & Param_ToStr("LIMT ", KDS8_ONOFF_ToStr(.LIMT))     'As Byte 'Begrenzung           'ON; OFF;
        s = s & Param_ToStr("KBD  ", SemiNib_ToStr(.KBD))         'As Byte 'Tastatur - Abtastung '0; 1; 2; 3;
    End With
    WFRM1_ToStr = s
End Function
Function WFRM2_TYP_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 1, 2: s = CStr(value)
    Case 3:    s = "XMOD"
    Case Else: s = "&H" & Hex$(value)
    End Select
    WFRM2_TYP_ToStr = s
End Function
Function WFRM2_ToStr(aWFRM2 As WFRM2) As String
    Dim s As String
    With aWFRM2
        s = s & Param_ToStr("TYP  ", WFRM2_TYP_ToStr(.TYP))       'As Byte 'Wellenformart        '1; 2; XMOD;  1: Sägezahn; 2: Rechteck; XMOD: osc2 moduliert osc1;
        s = s & Param_ToStr("SPCT ", Unsigned_3Bit1_ToStr(.SPCT)) 'As Byte 'Spektrum             '1; . . .; 8;
        s = s & Param_ToStr("RING ", SemiNib_ToStr(.RING))        'As Byte 'Glocken - Modulation '0; 1; 2; 3;
        s = s & Param_ToStr("LIMT ", KDS8_ONOFF_ToStr(.LIMT))     'As Byte 'Begrenzung           'ON; OFF;
        s = s & Param_ToStr("KBD  ", SemiNib_ToStr(.KBD))         'As Byte 'Tastatur - Abtastung '0; 1; 2; 3;
    End With
    WFRM2_ToStr = s
End Function

Function T_EG_ToStr(aT_EG As T_EG) As String
    Dim s As String: 's = "T_EG"
    With aT_EG
        s = s & Param_ToStr("TIMB ", Unsigned_7Bit99_ToStr(.TIMB)) 'As Byte 'Klangfarbe                     '0; . . .; 99;
        s = s & Param_ToStr("INT  ", Nibble_ToStr(.INT))           'As Byte 'Intensität                     '0; . . .; 15;
        s = s & Param_ToStr("ATK  ", Unsigned_5Bit_ToStr(.ATK))    'As Byte 'ATTACK,   Einschwingungszeit   '0; . . .; 31;
        s = s & Param_ToStr("DEC  ", Unsigned_5Bit_ToStr(.DEC))    'As Byte 'DECAY,    Ausschwingungszeit   '0; . . .; 31;
        s = s & Param_ToStr("SUS  ", Nibble_ToStr(.SUS))           'As Byte 'SUSTAIN,  Haltepunkt           '0; . . .; 15;
        s = s & Param_ToStr("REL  ", Nibble_ToStr(.REL))           'As Byte 'Release,  Abklingen            '0; . . .; 15;
        s = s & Param_ToStr("KBD  ", SemiNib_ToStr(.KBD))          'As Byte 'KEYBOARD, Tastatur - Abtastung '0; 1; 2; 3;
    End With
    T_EG_ToStr = s
End Function

Function A_EG_ToStr(aA_EG As A_EG) As String
    Dim s As String: 's = "A_EG"
    With aA_EG
        s = s & Param_ToStr("LEVL ", Unsigned_6Bit_ToStr(.LEVL)) 'LEVL As Byte '1 'Level, Pegel                   '0; . . .; 63;
        s = s & Param_ToStr("ATK  ", Unsigned_5Bit_ToStr(.ATK))  'ATK  As Byte '1 'ATTACK,   Einschwingungszeit   '0; . . .; 31;
        s = s & Param_ToStr("DEC  ", Unsigned_5Bit_ToStr(.DEC))  'DEC  As Byte '1 'DECAY,    Ausschwingungszeit   '0; . . .; 31;
        s = s & Param_ToStr("SUS  ", Nibble_ToStr(.SUS))         'SUS  As Byte '1 'SUSTAIN,  Haltepegel           '0; . . .; 15;
        s = s & Param_ToStr("REL  ", Nibble_ToStr(.REL))         'REL  As Byte '1 'Release,  Abklingen            '0; . . .; 15;
        s = s & Param_ToStr("KBD  ", SemiNib_ToStr(.KBD))        'KBD  As Byte '1 'KEYBOARD, Tastatur - Abtastung '0; 1; 2; 3;
    End With
    A_EG_ToStr = s
End Function
Public Function KDS8_MG_WF_ToStr(e As KDS8_MG_WF) As String
    Dim s As String
    Select Case e
    Case kds8_MG_TRI:  s = "TRI"  ' Triangle
    Case kds8_MG_SAW:  s = "SAW"  ' Sawtooth
    Case kds8_MG_SQUR: s = "SQUR" ' Square
    Case kds8_MG_S_H:  s = "S/H"  ' Sample&Hold
    Case Else:         s = "&H" & Hex$(e)
    End Select
    KDS8_MG_WF_ToStr = s
End Function
Public Function KDS8_MG_TASEL_ToStr(e As KDS8_MG_TASEL) As String
    Dim s As String
    Select Case e
    Case MGTA_OFF: s = "OFF"
    Case MGTA_1:   s = "1"
    Case MGTA_2:   s = "2"
    Case MGTA_1_2: s = "1+2"
    Case Else:     s = "&H" & Hex(e)
    End Select
    KDS8_MG_TASEL_ToStr = s
End Function

Function MG_ToStr(aMG As MG) As String
    Dim s As String
    With aMG
        s = s & Param_ToStr("WF   ", KDS8_MG_WF_ToStr(.WF))      'As KDS8_MG_WF    'WAVEFORM, Wellenform              'TRI; SAW; SQUR; S/H
        s = s & Param_ToStr("FREQ ", Unsigned_6Bit_ToStr(.FREQ)) 'As Byte          'FREQUENCE                         '0; . . .; 63;
        s = s & Param_ToStr("DLY  ", Unsigned_5Bit_ToStr(.DLY))  'As Byte          'DELAY                             '0; . . .; 31;
        s = s & Param_ToStr("PTCH ", Unsigned_6Bit_ToStr(.PTCH)) 'As Byte          'PITCH                             '0; . . .; 63;
        s = s & Param_ToStr("T/A  ", Unsigned_6Bit_ToStr(.T_A))  'As Byte          'TIMBRE/AMPLITUDE, Klangfarbe      '0; . . .; 63;
        s = s & Param_ToStr("TSEL ", KDS8_MG_TASEL_ToStr(.TSEL)) 'As KDS8_MG_TASEL 'TIMBRE SELECT, Klangfarbenwahl    'OFF, 1, 2, 1+2
        s = s & Param_ToStr("ASEL ", KDS8_MG_TASEL_ToStr(.ASEL)) 'As KDS8_MG_TASEL 'AMPLITUDE SELECT, Amplitudenwahl  'OFF, 1, 2, 1+2
    End With
    MG_ToStr = s
End Function
Public Function KDS8_PORTA_MODE_ToStr(ByVal e As KDS8_PORTA_MODE) As String
    Dim s As String
    Select Case e
    Case KDS8_PORTA_MODE.kds8_PortaMode_1: s = "1"
    Case KDS8_PORTA_MODE.kds8_PortaMode_2: s = "2"
    Case Else:             s = "&H" & Hex$(e)
    End Select
    KDS8_PORTA_MODE_ToStr = s
End Function

Function PORTAMENTO_ToStr(aPORTAMENTO As PORTAMENTO) As String
    Dim s As String
    With aPORTAMENTO
        s = s & Param_ToStr("MODE ", KDS8_PORTA_MODE_ToStr(.MODE)) 'MODE As KDS8_PORTA_MODE 'Modus '1; 2;
        s = s & Param_ToStr("TIME ", Unsigned_6Bit_ToStr(.Time))   'Time As Byte            'Zeit  '0; . . .; 63;
    End With
    PORTAMENTO_ToStr = s
End Function

Function JOYSTICK_ToStr(aJOYSTICK As JOYSTICK) As String
    Dim s As String
    With aJOYSTICK
        s = s & Param_ToStr("BEND:PITCH ", Signed_5Bit12_ToStr(.BEND_PITCH)) 'As Integer 'Biegung: Tonhöhe            '-12; . . .; +12;
        s = s & Param_ToStr("TIMB       ", SemiNib_ToStr(.BEND_TIMB))        'As Byte    'Biegung: Klangfarbe         '0; 1; 2; 3;
        s = s & Param_ToStr("MOD:SPEED  ", SemiNib_ToStr(.MOD_SPEED))        'As Byte    'Modulation: Geschwindigkeit '0; 1; 2; 3;
    End With
    JOYSTICK_ToStr = s
End Function

Function VELOCITY_ToStr(aVELOCITY As VELOCITY) As String
    Dim s As String
    With aVELOCITY
        s = s & Param_ToStr("TEG1", Unsigned_3Bit_ToStr(.TEG1)) 'As Byte 'Klangfarben EG1 '0; . . .; 7;
        s = s & Param_ToStr("TEG2", Unsigned_3Bit_ToStr(.TEG2)) 'As Byte 'Klangfarben EG2 '0; . . .; 7;
        s = s & Param_ToStr("AEG1", Unsigned_3Bit_ToStr(.AEG1)) 'As Byte 'Amplituden  EG1 '0; . . .; 7;
        s = s & Param_ToStr("AEG2", Unsigned_3Bit_ToStr(.AEG2)) 'As Byte 'Amplituden  EG2 '0; . . .; 7;
    End With
    VELOCITY_ToStr = s
End Function

Function AFT_TOUCH_ToStr(aAFT_TOUCH As AFT_TOUCH) As String
    Dim s As String: '"AFT_TOUCH"
    With aAFT_TOUCH
        s = s & Param_ToStr("PMG ", Unsigned_3Bit_ToStr(.PMG))  'As Byte 'PITCH MODULATION GENERATOR, Tonhöhen-Modulationserzeuger '0; . . .; 7;
        s = s & Param_ToStr("TIMB", Unsigned_3Bit_ToStr(.TIMB)) 'As Byte 'TIMBRE, Klangfarbe                '0; . . .; 7;
        s = s & Param_ToStr("AMP1", Unsigned_3Bit_ToStr(.AMP1)) 'As Byte 'OSC 1-Amplitude                   '0; . . .; 7;
        s = s & Param_ToStr("AMP2", Unsigned_3Bit_ToStr(.AMP2)) 'As Byte 'OSC 2-Amplitude                   '0; . . .; 7;
    End With
    AFT_TOUCH_ToStr = s
End Function

Public Function KDS8_ASSIGNMODE_ToStr(ByVal e As KDS8_ASSIGNMODE) As String
    Dim s As String
    Select Case e
    Case KDS8_ASSIGNMODE.kds8_POLY:      s = "POLY"
    Case KDS8_ASSIGNMODE.kds8_UNISON: s = "UNISON"
    Case Else: s = "&H" & Hex(e)
    End Select
    KDS8_ASSIGNMODE_ToStr = s
End Function
Public Function KDS8_UNISON_TRIG_ToStr(ByVal e As KDS8_UNISON_TRIG) As String
    Dim s As String
    Select Case e
    Case kds8_SINGLE: s = "SINGLE"
    Case kds8_MULTI:  s = "MULTI"
    Case Else: s = "&H" & Hex(e)
    End Select
    KDS8_UNISON_TRIG_ToStr = s
End Function
Function ASSIGNMODE_ToStr(aASSIGNMODE As ASSIGNMODE) As String
    Dim s As String: '"ASSIGNMODE"
    With aASSIGNMODE
        s = s & Param_ToStr("MODE", KDS8_ASSIGNMODE_ToStr(.MODE))      'As KDS8_ASSIGNMODE  'POLY; UNISON;  '
        If .MODE = KDS8_ASSIGNMODE.kds8_UNISON Then
            s = s & Param_ToStr("TRIG", KDS8_UNISON_TRIG_ToStr(.TRIG)) 'As KDS8_UNISON_TRIG 'SINGLE; MULTI; '
            s = s & Param_ToStr("DETUNE", SemiNib_ToStr(.DETUNE))      'As Byte             ' 0; 1; 2; 3;   'Die Funktion Detune erlaubt es die Tonhöhen der acht Stimmen, die beim Spielen einer einzelnen Note im Modus UNISON leicht zu verstimmen. Dadurch entsteht ein reicher Chor-Effekt mit einstellbarem Bereich.
        End If
    End With
    ASSIGNMODE_ToStr = s
End Function

Function VOICENAME_ToStr(aVOICENAME As VOICENAME) As String
    Dim s As String: 's = "VOICENAME"
    With aVOICENAME
        s = s & Param_ToStr("NAME ", .name)
    End With
    VOICENAME_ToStr = s
End Function
Public Function KDS8_EffectType_ToStr(ByVal e As KDS8_EffectType) As String
    Dim s As String
    Select Case e
    Case kds8_MANUAL_DLY: s = "MANUAL DLY"
    Case kds8_LONG_DLY:   s = "LONG DLY"
    Case kds8_SHORT_DLY:  s = "SHORT DLY"
    Case kds8_DOUBLING:   s = "DOUBLING"
    Case kds8_FLANGER:    s = "FLANGER"
    Case kds8_CHORUS:     s = "CHORUS"
    Case Else: s = "&H" & Hex$(e)
    End Select
    KDS8_EffectType_ToStr = s
End Function

Function MULTI_EFFECT_ToStr(aMULTI_EFFECT As MULTI_EFFECT) As String
    Dim s As String
    With aMULTI_EFFECT
        Select Case .EffectType
        Case kds8_MANUAL_DLY
            s = s & Param_ToStr("TIME", MANU_TIME_ToStr(.TIME_MANU))
            s = s & Param_ToStr("FB", CStr(.FB))
            s = s & Param_ToStr("MFRQ", CStr(.MFRQ))
            s = s & Param_ToStr("MINT", CStr(.MINT))
            
            '---
            s = s & Param_ToStr("SPED", CStr(.SPED))
            s = s & Param_ToStr("DPTH", CStr(.DPTH))
            '---
            
            s = s & Param_ToStr("LEVEL", CStr(.Level))
            
        Case kds8_LONG_DLY
            s = s & Param_ToStr("TIME", MANU_TIME_ToStr(.TIME_MANU))
            s = s & Param_ToStr("FB", CStr(.FB))
            
            '---
            s = s & Param_ToStr("MFRQ", CStr(.MFRQ))
            s = s & Param_ToStr("MINT", CStr(.MINT))
            s = s & Param_ToStr("SPED", CStr(.SPED))
            s = s & Param_ToStr("DPTH", CStr(.DPTH))
            '---
            
            s = s & Param_ToStr("LEVEL", CStr(.Level))
        Case kds8_SHORT_DLY
            s = s & Param_ToStr("TIME", MANU_TIME_ToStr(.TIME_MANU))
            s = s & Param_ToStr("FB", CStr(.FB))
            
            '---
            s = s & Param_ToStr("MFRQ", CStr(.MFRQ))
            s = s & Param_ToStr("MINT", CStr(.MINT))
            s = s & Param_ToStr("SPED", CStr(.SPED))
            s = s & Param_ToStr("DPTH", CStr(.DPTH))
            '---
            
            s = s & Param_ToStr("LEVEL", CStr(.Level))
        Case kds8_DOUBLING
            s = s & Param_ToStr("TIME", MANU_TIME_ToStr(.TIME_MANU))
            
            '----
            s = s & Param_ToStr("FB", CStr(.FB))
            s = s & Param_ToStr("MFRQ", CStr(.MFRQ))
            s = s & Param_ToStr("MINT", CStr(.MINT))
            s = s & Param_ToStr("SPED", CStr(.SPED))
            s = s & Param_ToStr("DPTH", CStr(.DPTH))
            '----
            
            s = s & Param_ToStr("LEVEL", CStr(.Level))
        Case kds8_FLANGER
            s = s & Param_ToStr("MANU", MANU_TIME_ToStr(.TIME_MANU))
            s = s & Param_ToStr("FB", CStr(.FB))
            
            '---
            s = s & Param_ToStr("MFRQ", CStr(.MFRQ))
            s = s & Param_ToStr("MINT", CStr(.MINT))
            '---
            
            s = s & Param_ToStr("SPED", CStr(.SPED))
            s = s & Param_ToStr("DPTH", CStr(.DPTH))
            s = s & Param_ToStr("LEVEL", CStr(.Level))
        Case kds8_CHORUS
            s = s & Param_ToStr("MANU", MANU_TIME_ToStr(.TIME_MANU))
            
            '---
            s = s & Param_ToStr("FB", CStr(.FB))
            s = s & Param_ToStr("MFRQ", CStr(.MFRQ))
            s = s & Param_ToStr("MINT", CStr(.MINT))
            '---
            
            s = s & Param_ToStr("SPED", CStr(.SPED))
            s = s & Param_ToStr("DPTH", CStr(.DPTH))
            s = s & Param_ToStr("LEVEL", CStr(.Level))
        End Select
    End With
    MULTI_EFFECT_ToStr = s
End Function

'Public Function MANUAL_DELAY_ToStr(aMANUAL_DELAY As MANUAL_DELAY) As String
'    Dim s As String: s = "MANUAL_DELAY"
'    With aMANUAL_DELAY
''        Time  As Byte    'Verzögerungszeit    0.04; . . .; 850; ms
''        FB    As Integer 'FEEDBACK Rückkopplung   -15; . . . ; +15;
''        MFRQ  As Byte    'MODULATION FREQUENCE Modulationsfrequenz    0; . . .; 31;
''        MINT  As Byte    'MODULATION INTENSITY Modulationsintensität  0; . . .; 15;
''        Level As Byte    'Pegel   0; . . .; 31;
'
'        s = s & Param_ToStr("TIME  ", CStr(.Time))
'        s = s & Param_ToStr("FB    ", CStr(.FB))
'        s = s & Param_ToStr("MFRQ  ", CStr(.MFRQ))
'        s = s & Param_ToStr("MINT  ", CStr(.MINT))
'        s = s & Param_ToStr("LEVEL ", CStr(.Level))
'    End With
'    MANUAL_DELAY_ToStr = s
'End Function
'Public Function LONG_DELAY_ToStr(aLONG_DELAY As LONG_DELAY) As String
'    Dim s As String: s = "LONG_DELAY"
'    With aLONG_DELAY
''        Time  As Byte    'Verzögerungszeit    105; . . . ; 720; ms
''        FB    As Integer 'FEEDBACK Rückkopplung   -15; . . .; +15;
''        Level As Byte    'Pegel   0; . . .; 31;
'
'        s = s & Param_ToStr("TIME  ", CStr(.Time))
'        s = s & Param_ToStr("FB    ", CStr(.FB))
'        s = s & Param_ToStr("LEVEL ", CStr(.Level))
'    End With
'    LONG_DELAY_ToStr = s
'End Function
'Public Function SHORT_DELAY_ToStr(aSHORT_DELAY As SHORT_DELAY) As String
'    Dim s As String: s = "SHORT_DELAY"
'    With aSHORT_DELAY
''        Time  As Byte    'Verzögerungszeit    20; . . .; 88; ms
''        FB    As Integer 'FEEDBACK Rückkopplung   -15; . . .; +15;
''        Level As Byte    'Pegel   0; . . .; 31;
'
'        s = s & Param_ToStr("TIME  ", CStr(.Time))
'        s = s & Param_ToStr("FB    ", CStr(.FB))
'        s = s & Param_ToStr("LEVEL ", CStr(.Level))
'    End With
'    SHORT_DELAY_ToStr = s
'End Function
'Public Function DOUBLING_ToStr(aDOUBLING As DOUBLING) As String
'    Dim s As String: s = "DOUBLING"
'    With aDOUBLING
''        Time  As Byte    'Verzögerungszeit
''        Level As Byte    'Pegel   0; . . .; 31;
'
'        s = s & Param_ToStr("TIME  ", CStr(.Time))
'        s = s & Param_ToStr("LEVEL ", CStr(.Level))
'    End With
'    DOUBLING_ToStr = s
'End Function
'Public Function FLANGER_ToStr(aFLANGER As FLANGER) As String
'    Dim s As String: s = "FLANGER"
'    With aFLANGER
''        MANU  As Byte    'MANUAL Manuell  0.04; . . .; 5.5; ms    Bestimmt die verzögerungszeit zwischen dem Direkt- und dem Flanger-Signal.
''        FB    As Integer 'FEEDBACK Rückkopplung   -15; . . .; +15;    Bestimmt den Wert, mit dem das Flanger-Signal mit sich selbst zurückgekoppelt wird und bestimmt die Intensität des Effekts. Negative Einstellungen produzieren ein phasen-umgekehrtes Flanging, was in einen klareren, helleren Sound resultieren kann.
''        SPED  As Byte    'SPEED Flanger-Modulationsgeschwindigkeit    0; . . .; 24;   Bestimmt die Intensität des Flanger-Effektes, der von einer langsamen, ruhigen Modulation bis zu einer schnellen, extremen Einstellung reicht.
''        DPTH  As Byte    'DEPTH Flanger-Modulationstiefe  0; . . .; 31;   Bestimmt die Flanger-Modulationstiefe, die von Null (kein Effekt) bi zu einem hochmodulierten Sound reicht.
''        Level As Byte    'Pegel   0; . . .; 15;   Bestimmt den Gesamtpegel des Flanger-Signals.
'
'        s = s & Param_ToStr("MANU  ", CStr(.MANU))
'        s = s & Param_ToStr("FB    ", CStr(.FB))
'        s = s & Param_ToStr("SPED  ", CStr(.SPED))
'        s = s & Param_ToStr("DPTH  ", CStr(.DPTH))
'        s = s & Param_ToStr("Level ", CStr(.Level))
'    End With
'    FLANGER_ToStr = s
'End Function
'Public Function CHORUS_ToStr(aCHORUS As CHORUS) As String
'    Dim s As String: s = "CHORUS"
'    With aCHORUS
''        MANU  As Byte   'MANUAL  5.0; . . .; 32; ms  Bestimmt die Verzögerungszeit zwischen dem Direkt- und dem Chorus-Signal
''        SPED  As Byte   'SPEED   0; . . .; 31;   Bestimmt die Intensität des Chorus-Effektes, der von einer langsamen, ruhigen Modulation bis zu einer schnellen, extremen Einstellung reicht.
''        DPTH  As Byte   'DEPTH   0; . . .; 31;   Bestimmt die Chorus-Modulationstiefe, die von Null (kein Effekt) bis zu einem hochmodulierten Sound reicht.
''        Level As Byte   'Pegel   0; . . .; 31;   Bestimmt den Gesamtpegl des Chorus-Signals.
'
'        s = s & Param_ToStr("MANU  ", CStr(.MANU))
'        s = s & Param_ToStr("SPED  ", CStr(.SPED))
'        s = s & Param_ToStr("DPTH  ", CStr(.DPTH))
'        s = s & Param_ToStr("Level ", CStr(.Level))
'    End With
'    CHORUS_ToStr = s
'End Function

Function MANU_TIME_ToStr(ByVal mt As Byte) As String
    Dim s As String
    Select Case mt
    Case 0:    s = "0.04"
    Case 1:    s = "0.12"
    Case 2:    s = "0.28"
    Case 3:    s = "0.36"
    Case 4:    s = "0.52"
    Case 5:    s = "0.76"
    Case 6:    s = "1.00"
    Case 7:    s = "1.50"
    Case 8:    s = "2.00"
    Case 9:    s = "2.50"
    Case 10:   s = "3.00"
    Case 11:   s = "3.50"
    Case 12:   s = "4.00"
    Case 13:   s = "4.50"
    Case 14:   s = "5.00"
    Case 15:   s = "5.50"
    Case 16:   s = "6.00"
    Case 17:   s = "6.50"
    Case 18:   s = "7.00"
    Case 19:   s = "7.50"
    Case 20:   s = "8.00"
    Case 21:   s = "8.50"
    Case 22:   s = "9.00"
    Case 23:   s = "9.50"
    Case 24:   s = "10.0"
    Case 25:   s = "11.0"
    Case 26:   s = "12.0"
    Case 27:   s = "13.0"
    Case 28:   s = "14.0"
    Case 29:   s = "15.0"
    Case 30:   s = "16.0"
    Case 31:   s = "17.0"
    Case 32:   s = "18.0"
    Case 33:   s = "19.0"
    Case 34:   s = "20.0"
    Case 35:   s = "21.0"
    Case 36:   s = "22.0"
    Case 37:   s = "23.0"
    Case 38:   s = "24.0"
    Case 39:   s = "25.0"
    Case 40:   s = "26.0"
    Case 41:   s = "27.0"
    Case 42:   s = "28.0"
    Case 43:   s = "29.0"
    Case 44:   s = "30.0"
    Case 45:   s = "32.0"
    Case 46:   s = "34.0"
    Case 47:   s = "36.0"
    Case 48:   s = "38.0"
    Case 49:   s = "40.0"
    Case 50:   s = "42.0"
    Case 51:   s = "44.0"
    Case 52:   s = "46.0"
    Case 53:   s = "48.0"
    Case 54:   s = "50.0"
    Case 55:   s = "52.0"
    Case 56:   s = "55.0"
    Case 57:   s = "58.0"
    Case 58:   s = "60.0"
    Case 59:   s = "62.0"
    Case 60:   s = "65.0"
    Case 61:   s = "68.0"
    Case 62:   s = "70.0"
    Case 63:   s = "72.0"
    Case 64:   s = "75.0"
    Case 65:   s = "78.0"
    Case 66:   s = "80.0"
    Case 67:   s = "82.0"
    Case 68:   s = "85.0"
    Case 69:   s = "88.0"
    Case 70:   s = "90.0"
    Case 71:   s = "92.0"
    Case 72:   s = "95.0"
    Case 73:   s = "98.0"
    Case 74:   s = "100."
    Case 75:   s = "105."
    Case 76:   s = "110."
    Case 77:   s = "115."
    Case 78:   s = "120."
    Case 79:   s = "125."
    Case 80:   s = "130."
    Case 81:   s = "135."
    Case 82:   s = "140."
    Case 83:   s = "145."
    Case 84:   s = "150."
    Case 85:   s = "160."
    Case 86:   s = "170."
    Case 87:   s = "180."
    Case 88:   s = "190."
    Case 89:   s = "200."
    Case 90:   s = "210."
    Case 91:   s = "220."
    Case 92:   s = "230."
    Case 93:   s = "240."
    Case 94:   s = "250."
    Case 95:   s = "260."
    Case 96:   s = "270."
    Case 97:   s = "280."
    Case 98:   s = "290."
    Case 99:   s = "300."
    Case 100:  s = "310."
    Case 101:  s = "320."
    Case 102:  s = "330."
    Case 103:  s = "340."
    Case 104:  s = "350."
    Case 105:  s = "360."
    Case 106:  s = "370."
    Case 107:  s = "380."
    Case 108:  s = "390."
    Case 109:  s = "400."
    Case 110:  s = "420."
    Case 111:  s = "450."
    Case 112:  s = "480."
    Case 113:  s = "500."
    Case 114:  s = "520."
    Case 115:  s = "550."
    Case 116:  s = "580."
    Case 117:  s = "600."
    Case 118:  s = "620."
    Case 119:  s = "650."
    Case 120:  s = "680."
    Case 121:  s = "700."
    Case 122:  s = "720."
    Case 123:  s = "750."
    Case 124:  s = "780."
    Case 125:  s = "800."
    Case 126:  s = "820."
    Case 127:  s = "850."
    End Select
    MANU_TIME_ToStr = s
End Function
