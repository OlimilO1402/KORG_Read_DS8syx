Attribute VB_Name = "MKDS8Reader"
Option Explicit
Public Const MIDI_SysEX  As Byte = &HF0
Public Const ID_KORG     As Byte = &H42
Public Const ID_KORG_DS8 As Byte = &H13

Public KORG_DS8_PROGS As ListOf_KORG_DS8_PROG

Public Sub List_Add_KORG_DS8_PROG(aKORG_DS8_PROG As KORG_DS8_PROG)
    With KORG_DS8_PROGS
        If .Count = 0 Then
            ReDim .Arr(0 To 3)
        Else
            Dim u As Long: u = UBound(.Arr)
            If .Count > u Then
                ReDim Preserve .Arr(0 To 2 * u - 1)
            End If
        End If
        .Arr(.Count) = aKORG_DS8_PROG
        .Count = .Count + 1
    End With
End Sub
Public Sub List_ToComboBox(aCmb As ComboBox)
    Dim i As Long
    With KORG_DS8_PROGS
        For i = 0 To .Count - 1
            aCmb.AddItem .Arr(i).VOICENAME.name
        Next
    End With
End Sub
Public Function List_Get_Item(ByVal aName As String) As KORG_DS8_PROG
    Dim i As Long
    With KORG_DS8_PROGS
        For i = 0 To .Count - 1
            If .Arr(i).VOICENAME.name = aName Then
                List_Get_Item = .Arr(i)
                Exit Function
            End If
        Next
    End With
End Function
Public Function ReadPITCH(ByVal sr As KDS8SyxReader) As PITCH
    With ReadPITCH
        .OSC1 = sr.ReadByte 'As Byte 'Oscillator 1 '0.5; . . . ; 15;   0.5 (16-Fuss-Orgel); 1 (8-Fuss-Orgel); 2 (4-Fuss-Orgel); . . .
        .OSC2 = sr.ReadByte 'As Byte 'Oscillator 1 '0.5; . . . ; 15;   0.5 (16-Fuss-Orgel); 1 (8-Fuss-Orgel); 2 (4-Fuss-Orgel); . . .
        .DTN = sr.ReadByte  'As Byte 'Detune       '0; 1; 2; 3
    End With
End Function

Public Function ReadPITCH_EG(ByVal sr As KDS8SyxReader) As PITCH_EG
    With ReadPITCH_EG
        .STL = sr.ReadSigned7Bit 'As Integer 'Anfangspegel         '-63; . . . ; +63; 'Die Tonhöhe bei der die Klangfarbe beginnt jedesmal wen eine Note gespielt wird
        .ATK = sr.ReadByte       'As Byte    'Einschwingungszeit   '  0; . . . ;  63; 'Die Zeit in der sich die Tonhöhe vom Anfangspegel zum einschwingungspegel ändert
        .ATL = sr.ReadSigned7Bit 'As Integer 'Einschwingungspegel  '-63; . . . ; +63; 'Der Spitzenpegelwert der Tonhöhe
        .DEC = sr.ReadByte       'As Byte    'Ausschwingungszeit   '  0; . . . ;  63; 'Die Zeit in der die Tonhöhe vom Einschwingungspegel zum Normalpegel zurückkehrt, währen die Taste gedrücktgehalten bleibt.
        .REL = sr.ReadByte       'As Byte    'Abklingung           '  0; . . . ;  63; 'Die Zeit in der die Tonhöht zum Abklingungspegel wechselt, nachdem die Taste losgelassen wurde.
        .RLL = sr.ReadSigned7Bit 'As Integer 'Abklingungspegel     '-63; . . . ; +63; 'Der Pegel zu dem die Tonhöhe wechselt, nachdem die Taste losgelassen wurde.
    End With
End Function
Public Function ReadWFRM1(ByVal sr As KDS8SyxReader) As WFRM1
    With ReadWFRM1
        Dim b As Byte
        b = sr.ReadByte
        .TYP = Lo2Bit8(b) + 1  'As Byte       'Wellenformart        '1; 2; 3; 4;  1: Sägezahn; 2: Rechteck; 3: Heller Sägezahn; 4: Helles Rechteck;
        .SPCT = Hi6Bit8(b) + 1 'As Byte       'Spektrum             '1; . . .; 8;
        b = sr.ReadByte
        .RING = LoSemiNib(b)     'As Byte       'Glocken - Modulation '0; 1; 2; 3;
        .LIMT = HiSemiNib(b)     'As KDS8_ONOFF 'Begrenzung           'ON; OFF;
        b = sr.ReadByte
        .KBD = b                 'As Byte       'Tastatur - Abtastung '0; 1; 2; 3;
    End With
End Function
Public Function ReadWFRM2(ByVal sr As KDS8SyxReader) As WFRM2
    With ReadWFRM2
        Dim b As Byte
        b = sr.ReadByte
        .TYP = Lo2Bit8(b) + 1  'As Byte       'Wellenformart        '1; 2; XMOD;  1: Sägezahn; 2: Rechteck; XMOD: osc2 moduliert osc1;
        .SPCT = Hi6Bit8(b) + 1 'As Byte       'Spektrum             '1; . . .; 8;
        b = sr.ReadByte
        .RING = LoSemiNib(b)     'As Byte       'Glocken - Modulation '0; 1; 2; 3;
        .LIMT = HiSemiNib(b)     'As KDS8_ONOFF 'Begrenzung           'ON; OFF;
        b = sr.ReadByte
        .KBD = b                 'As Byte       'Tastatur - Abtastung '0; 1; 2; 3;
    End With
End Function

'Public Function ReadT_EG(ByVal sr As KDS8SyxReader) As T_EG
'    Debug.Print "T_EG"
'    With ReadT_EG
'        Dim b As Byte
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .TIMB = b 'sr.ReadByte 'As Byte 'Klangfarbe                      '0; . . .; 99;
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .INT = b 'sr.ReadByte  'As Byte 'Intensität                      '0; . . .; 15;
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .ATK = LoNib(b)     'As Byte 'ATTACK,   'Einschwingungszeit   '0; . . .; 31;
'        .DEC = HiNib(b)     'As Byte 'DECAY,    'Ausschwingungszeit   '0; . . .; 31;
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .SUS = LoSemiNib(b) 'As Byte 'SUSTAIN,  'Haltepunkt           '0; . . .; 15;
'        .REL = HiSemiNib(b) 'As Byte 'Release,  'Abklingen            '0; . . .; 15;
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .KBD = b 'sr.ReadByte  'As Byte 'KEYBOARD, 'Tastatur - Abtastung '0; 1; 2; 3;
'    End With
'End Function
Public Function ReadT_EG12(ByVal sr As KDS8SyxReader) As T_EG12
    Debug.Print "T_EG1"
    Dim b As Byte
    With ReadT_EG12
        With .OSC1
            'Dim b As Byte
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .TIMB = b 'sr.ReadByte 'As Byte 'Klangfarbe                      '0; . . .; 99;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .INT = b 'sr.ReadByte  'As Byte 'Intensität                      '0; . . .; 15;
            
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .ATK = Lo5Bit8(b)     'As Byte 'ATTACK,   'Einschwingungszeit   '0; . . .; 31;
            .KBD = Hi2Bit8(b) 'sr.ReadByte  'As Byte 'KEYBOARD, 'Tastatur - Abtastung '0; 1; 2; 3;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .DEC = b             'As Byte 'DECAY,    'Ausschwingungszeit   '0; . . .; 31;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .SUS = LoNib(b) 'As Byte 'SUSTAIN,  'Haltepunkt           '0; . . .; 15;
            .REL = HiNib(b) 'As Byte 'Release,  'Abklingen            '0; . . .; 15;
            'b = sr.ReadByte:        Debug.Print Hex(b)
        End With
    Debug.Print "----"
    Debug.Print "T_EG2"
        With .OSC2
            'Dim b As Byte
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .TIMB = b 'sr.ReadByte 'As Byte 'Klangfarbe                      '0; . . .; 99;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .INT = b 'sr.ReadByte  'As Byte 'Intensität                      '0; . . .; 15;
            
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .ATK = Lo5Bit8(b)     'As Byte 'ATTACK,   'Einschwingungszeit   '0; . . .; 31;
            
' Ohjemine, also KBD von T_EG2 könnte auf dem odd8-byte aufgesetzt sein.
' also muss man in der lage sein dieses byte auszulesen, noch bevor es weggeschmissen wird.
' Entweder, oder
' * man ändert die Automatik so, dass genau dieses byte nicht weggeschmissen wird
' * man schaltet die Automatik zwischenzeitlich aus
' * die SyxReader-Klasse muss das Byte stets zur Verfügung stellen
'   d.h. das odd-Byte muss voraus-ausgelesen werden und zwischengespeichert werden.
            
            '.KBD = Hi2Bit8(b) 'sr.ReadByte  'As Byte 'KEYBOARD, 'Tastatur - Abtastung '0; 1; 2; 3;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .DEC = b             'As Byte 'DECAY,    'Ausschwingungszeit   '0; . . .; 31;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .SUS = LoNib(b) 'As Byte 'SUSTAIN,  'Haltepunkt           '0; . . .; 15;
            .REL = HiNib(b) 'As Byte 'Release,  'Abklingen            '0; . . .; 15;
            'b = sr.ReadByte:        Debug.Print Hex(b)
            b = sr.LastSkippedByte
            .KBD = HiNib(b)
        End With
    End With
    Debug.Print "----"
End Function
'
'Tja wie ist das
'man müßte jetzt hergehen und alle Parameter zu Null setzen
'die Syx-Datei runterladen und nachschauen
' * ob alle Paremter wirklich 0 sind,
' * was mit den 8-er Füllbytes passiert ist
'dann müßte man, successive alle parameter über ihren gesamten Wertebereich variieren
'
'Public Function ReadA_EG(ByVal sr As KDS8SyxReader) As A_EG
''    LEVL As Byte '1 'Level, Pegel                   '0; . . .; 63;
''    'MyEG As EG
''    ATK  As Byte '1 'ATTACK,   Einschwingungszeit   '0; . . .; 31;
''    DEC  As Byte '1 'DECAY,    Ausschwingungszeit   '0; . . .; 31;
''    SUS  As Byte '1 'SUSTAIN,  Haltepegel           '0; . . .; 15;
''    REL  As Byte '1 'Release,  Abklingen            '0; . . .; 15;
''    KBD  As Byte '1 'KEYBOARD, Tastatur - Abtastung '0; 1; 2; 3;
'
'    Debug.Print "A_EG"
'    With ReadA_EG
'        Dim b As Byte
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .LEVL = b 'sr.ReadByte 'As Byte 'Level, Pegel                   '0; . . .; 63;
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .ATK = b 'sr.ReadByte 'LoNib(b)     'As Byte 'ATTACK,   Einschwingungszeit '0; . . .; 31;
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .DEC = b 'sr.ReadByte  'As Byte 'DECAY,    Ausschwingungszeit   '0; . . .; 31;
'        b = sr.ReadByte:        Debug.Print Hex(b)
'        .SUS = LoNib(b)     'As Byte 'SUSTAIN,  Haltepunkt           '0; . . .; 15;
'        .REL = HiNib(b)     'As Byte 'Release,  Abklingen            '0; . . .; 15;
'        .KBD = b            'As Byte 'KEYBOARD, Tastatur-Abtastung '0; 1; 2; 3;
'        Debug.Print "--"
'    End With
'End Function
Public Function ReadA_EG12(ByVal sr As KDS8SyxReader) As A_EG12
    Debug.Print "A_EG1"
    Dim b As Byte
    With ReadA_EG12
        With .OSC1
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .LEVL = b 'sr.ReadByte 'As Byte 'Level, Pegel                   '0; . . .; 63;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .ATK = Lo5Bit8(b) 'sr.ReadByte 'LoNib(b)     'As Byte 'ATTACK,   Einschwingungszeit '0; . . .; 31;
            .KBD = Hi3Bit8(b)            'As Byte 'KEYBOARD, Tastatur-Abtastung  '0; 1; 2; 3;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .DEC = b 'sr.ReadByte  'As Byte 'DECAY,    Ausschwingungszeit '0; . . .; 31;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .SUS = LoNib(b)     'As Byte 'SUSTAIN,  Haltepunkt '0; . . .; 15;
            .REL = HiNib(b)     'As Byte 'Release,  Abklingen  '0; . . .; 15;
            'Debug.Print "--"
        End With
    Debug.Print "----"
    Debug.Print "A_EG2"
        With .OSC2
            'Dim b As Byte
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .LEVL = b 'sr.ReadByte 'As Byte 'Level, Pegel                   '0; . . .; 63;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .ATK = Lo5Bit8(b) 'sr.ReadByte 'LoNib(b)     'As Byte 'ATTACK,   Einschwingungszeit '0; . . .; 31;
            .KBD = Hi3Bit8(b)               'As Byte 'KEYBOARD, Tastatur-Abtastung '0; 1; 2; 3;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .DEC = b 'sr.ReadByte  'As Byte 'DECAY,    Ausschwingungszeit '0; . . .; 31;
            b = sr.ReadByte:        'Debug.Print Hex(b)
            .SUS = LoNib(b)        'As Byte 'SUSTAIN,  Haltepunkt '0; . . .; 15;
            .REL = HiNib(b)        'As Byte 'Release,  Abklingen  '0; . . .; 15;
            'Debug.Print "--"
        End With
    End With
    Debug.Print "----"
End Function
Public Function ReadMG(ByVal sr As KDS8SyxReader) As MG 'MODULATION GENERATOR Modulationserzeuger
    With ReadMG
        .WF = sr.ReadByte   'As KDS8_MG_WF     'WAVEFORM, Wellenform             'TRI; SAW; SQUR; S/H
        .FREQ = sr.ReadByte 'As Byte           'FREQUENCE                        '0; . . .; 63;
        .DLY = sr.ReadByte  'As Byte           'DELAY                            '0; . . .; 31;
        .PTCH = sr.ReadByte 'As Byte           'PITCH                            '0; . . .; 63;
        .T_A = sr.ReadByte  'As Byte           'TIMBRE/AMPLITUDE, Klangfarbe     '0; . . .; 63;
        Dim b As Byte
        b = sr.ReadByte
        .TSEL = LoSemiNib(b) 'As KDS8_MG_TASEL 'TIMBRE SELECT, Klangfarbenwahl   'OFF, 1, 2, 1+2
        .ASEL = HiSemiNib(b) 'As KDS8_MG_TASEL 'AMPLITUDE SELECT, Amplitudenwahl 'OFF, 1, 2, 1+2
    End With
End Function

Public Function ReadPORTAMENTO(ByVal sr As KDS8SyxReader) As PORTAMENTO
'vielleicht ist hier nur 1 Byte
'MODE ist 0 (=1) und 1 (=2)
    With ReadPORTAMENTO
        Dim b As Byte
        b = sr.ReadByte
        .MODE = LoSemiNib(b) 'sr.ReadByte 'As Byte 'Modus  '1; 2;
        .Time = b 'HiNib(b) 'sr.ReadByte 'As Byte 'Zeit   '0; . . .; 63;
    End With
End Function

Public Function ReadJOYSTICK(ByVal sr As KDS8SyxReader) As JOYSTICK
    With ReadJOYSTICK
    '-12 bis 12 was soll denn das für ein Wert sein?
        .BEND_PITCH = sr.ReadByte 'As Integer 'Biegung: Tonhöhe            '-12; . . .; +12;
        .BEND_TIMB = sr.ReadByte  'As Byte    'Biegung: Klangfarbe         '0; 1; 2; 3;
        .MOD_SPEED = sr.ReadByte  'As Byte    'Modulation: Geschwindigkeit '0; 1; 2; 3;
    End With
End Function
Public Function ReadVELOCITY(ByVal sr As KDS8SyxReader) As VELOCITY
    With ReadVELOCITY
        Dim b As Byte
        b = sr.ReadByte
        'oder auch 3 bit?
        .TEG1 = b 'Lo3Bit(b) 'As Byte 'Klangfarben EG1                   '0; . . .; 7;
        b = sr.ReadByte
        .TEG2 = b 'Hi3Bit(b) 'As Byte 'Klangfarben EG2                   '0; . . .; 7;
        b = sr.ReadByte
        .AEG1 = b 'Lo3Bit(b) 'As Byte 'Amplituden EG1                    '0; . . .; 7;
        b = sr.ReadByte
        .AEG2 = b 'Hi3Bit(b) 'As Byte 'Amplituden EG2                    '0; . . .; 7;
    End With
End Function
Public Function ReadAFT_TOUCH(ByVal sr As KDS8SyxReader) As AFT_TOUCH
    With ReadAFT_TOUCH
        Dim b As Byte
        b = sr.ReadByte
        .PMG = Lo3Bit6(b)  'As Byte 'PITCH MODULATION GENERATOR, Tonhöhen-Modulationserzeuger '0; . . .; 7;
        .TIMB = Hi3Bit6(b) 'As Byte 'TIMBRE, Klangfarbe                '0; . . .; 7;
        'b = sr.ReadByte
        .AMP1 = Lo3Bit6(b) 'As Byte 'OSC 1-Amplitude                   '0; . . .; 7;
        .AMP2 = Hi3Bit6(b) 'As Byte 'OSC 2-Amplitude                   '0; . . .; 7;
    End With
End Function
Public Function ReadASSIGNMODE(ByVal sr As KDS8SyxReader) As ASSIGNMODE
    With ReadASSIGNMODE
        Dim b As Byte
        b = sr.ReadByte
        'MsgBox "&H" & Hex(b)
        .MODE = LoSemiNib(b)   'As KDS8_ASSIGNMODE  'POLY; UNISON;  '
        .TRIG = HiSemiNib(b)   'As KDS8_UNISON_TRIG 'SINGLE; MULTI; '
        .DETUNE = HiNib(b) 'As Byte             ' 0; 1; 2; 3;   'Die Funktion Detune erlaubt es die Tonhöhen der acht Stimmen, die beim Spielen einer einzelnen Note im Modus UNISON leicht zu verstimmen. Dadurch entsteht ein reicher Chor-Effekt mit einstellbarem Bereich.
        'b = sr.ReadByte
    End With
End Function
Public Function ReadVOICENAME(ByVal sr As KDS8SyxReader) As VOICENAME
'What about this:
'wenn es  9 Zeichen sind, dann 2 null dazu dann sind es 11 bytes
'wenn es 10 zeichen sind, dann 2 null dazu dann sind es 12 bytes
'ansonsten sind es immer maximal 10 bytes
'kommen zwei nullen dann ist der String zuende, d.h. der Name ist kürzer als 9 Zeichen
'es werden aber trotzdem mindestens 10 bytes gelesen
'
    With ReadVOICENAME
    'wie lang? nullterminiert?
        Dim b As Byte
        Dim s As String: 's = """"
        Dim s1 As String ': s1 = """"
        Dim i As Long
        Dim nn As Long
        'Debug.Print "---"
        'Do While i < 10
        For i = 0 To 9 '11
        'Do While b > 0
            'Get FNr, , b
            b = sr.ReadByte
            If b = 0 Then
                'Debug.Print b
                'nn = nn + 1
                's = s & " "
                s1 = s1 & " "
            Else
                s = s & Chr$(b + 32)
                s1 = s1 & Chr$(b + 32)
                'Debug.Print Chr$(b + 32)
                If Len(s) = 10 Then Exit For
                'i = i + 1
                'nn = 0
            End If
'            If nn = 2 Then
'                'den Dateizeiger um eins zurücksetzen
'                If i > 9 Then
'                    Dim dz As Long
'                    dz = Seek(FNr)
'                    'seek(fnr) = dz - 1
'                    Get FNr, dz - 2, b
'                Else
'                    For i = i To 9
'                        Get FNr, , b
'                    Next
'                End If
'                Exit Do
'            End If
        'Loop
        Next
        's = """" & s & """"
'        If i >= 11 And Len(s) < 10 Then
'            Debug.Print sr.FilePos 'Seek(FNr)
'            Debug.Print "SETBACK-1"
'            Dim dz As Long
'            dz = sr.FilePos 'Seek(FNr)
'            'seek(fnr) = dz - 1
'            'Get FNr, dz - 2, b
'            sr.SetBack 1
'            Debug.Print sr.FilePos 'Seek(FNr)
'        End If
        'Debug.Print "Seek(FNr): " & sr.FilePos 'Seek(FNr)
        'Debug.Print "---"
        .name = s1 '& """"
    End With
End Function
Public Function ReadMULTI_EFFECT(ByVal sr As KDS8SyxReader) As MULTI_EFFECT
    With ReadMULTI_EFFECT
        Dim b As Byte
        'Get FNr, , b:
        .EffectType = sr.ReadByte
        Select Case .EffectType
        Case kds8_MANUAL_DLY
            .TIME_MANU = sr.ReadByte
            .FB = sr.ReadSigned5Bit
            .MFRQ = sr.ReadByte
            .MINT = sr.ReadByte
            .Level = sr.ReadByte
            '.DPTH = sr.ReadByte
        Case kds8_LONG_DLY
            .TIME_MANU = sr.ReadByte
            .FB = sr.ReadSigned5Bit
            .Level = sr.ReadByte
            '.MFRQ = sr.ReadByte
            '.SPED = sr.ReadByte
            '.MINT = sr.ReadByte
            '.DPTH = sr.ReadByte
            b = sr.ReadByte
            b = sr.ReadByte
            
        Case kds8_SHORT_DLY
            .TIME_MANU = sr.ReadByte
            .FB = sr.ReadSigned5Bit
            .Level = sr.ReadByte
            b = sr.ReadByte
            b = sr.ReadByte
            '.MFRQ = sr.ReadByte
            '.SPED = sr.ReadByte
            '.MINT = sr.ReadByte
            '.DPTH = sr.ReadByte
        Case kds8_DOUBLING
            .TIME_MANU = sr.ReadByte
            .Level = sr.ReadByte
            '.FB = sr.ReadSigned5Bit
            b = sr.ReadByte
            b = sr.ReadByte
            b = sr.ReadByte
            '.MFRQ = sr.ReadByte
            '.SPED = sr.ReadByte
            '.MINT = sr.ReadByte
            '.DPTH = sr.ReadByte
            
        Case kds8_FLANGER
            .TIME_MANU = sr.ReadByte
            .FB = sr.ReadSigned5Bit
            .SPED = sr.ReadByte
            .DPTH = sr.ReadByte
            .Level = sr.ReadByte
            '.MFRQ = sr.ReadByte
            '.DPTH = sr.ReadByte
            
        Case kds8_CHORUS
            .TIME_MANU = sr.ReadByte
            .Level = sr.ReadByte
            .SPED = sr.ReadByte
            .DPTH = sr.ReadByte
            '.FB = sr.ReadSigned5Bit
            '.MFRQ = sr.ReadByte
            '.MINT = sr.ReadByte
            b = sr.ReadByte
        End Select
    End With
End Function
'Public Function ReadMANUAL_DELAY(ByVal FNr As Integer) As MANUAL_DELAY
'    With ReadMANUAL_DELAY
'        Get FNr, , .Time
'        Dim b As Byte
'        Get FNr, , b
'        .FB = b - 15
'        Get FNr, , .MFRQ
'        Get FNr, , .MINT
'        Get FNr, , b
'        Get FNr, , .Level
'        Get FNr, , b
'    End With
'End Function
'Public Function ReadLONG_DELAY(ByVal FNr As Integer) As LONG_DELAY
'    With ReadLONG_DELAY
'        Get FNr, , .Time
'        Dim b As Byte
'        Get FNr, , b
'        .FB = b - 15
'        Get FNr, , .Level
'    End With
'End Function
'Public Function ReadSHORT_DELAY(ByVal FNr As Integer) As SHORT_DELAY
'    With ReadSHORT_DELAY
'        Get FNr, , .Time
'        Dim b As Byte
'        Get FNr, , b
'        .FB = b - 15
'        Get FNr, , .Level
'    End With
'End Function
'Public Function ReadDOUBLING(ByVal FNr As Integer) As DOUBLING
'    With ReadDOUBLING
'        Get FNr, , .Time
'        Get FNr, , .Level
'        Dim b As Byte
'        Get FNr, , b
'        Get FNr, , b
'        Get FNr, , b
'        Get FNr, , b
'        Get FNr, , b
'    End With
'End Function
'Public Function ReadFLANGER(ByVal FNr As Integer) As FLANGER
'    With ReadFLANGER
'        Get FNr, , .MANU
'        Dim b As Byte
'        Get FNr, , b
'        .FB = b - 15
'        Get FNr, , .SPED
'        Get FNr, , .DPTH
'        Get FNr, , .Level
'    End With
'End Function
'Public Function ReadCHORUS(ByVal FNr As Integer) As CHORUS
'    With ReadCHORUS
'        Get FNr, , .MANU
'        Get FNr, , .SPED
'        Get FNr, , .DPTH
'        Get FNr, , .Level
'        Dim b As Byte
'        Get FNr, , b
'        Get FNr, , b
'        Get FNr, , b
'    End With
'End Function
'Public Function ReadMEEmpty(ByVal FNr As Integer)
'    Dim b As Byte
'    Get FNr, , b
'    Get FNr, , b
'    Get FNr, , b
'    Get FNr, , b
'    Get FNr, , b
'    Get FNr, , b
'    Get FNr, , b
'
'End Function
Public Function ReadKORGDS8_Prog(ByVal sr As KDS8SyxReader) As KORG_DS8_PROG
    Dim b As Byte ': Get FNr, , b
    'If b = 0 Then
    '    Debug.Print "OK start"
    'Else
    '    Debug.Print "Oh NO"
    'End If
    With ReadKORGDS8_Prog
        .PITCH = ReadPITCH(sr)
        .PITCH_EG = ReadPITCH_EG(sr)
        
        .OSC1_WFRM1 = ReadWFRM1(sr)
        .OSC2_WFRM2 = ReadWFRM2(sr)
        
        .TIMBRE_EG = ReadT_EG12(sr)
        
        '.OSC1_TIMBRE_EG = ReadT_EG(sr)
        '.OSC2_TIMBRE_EG = ReadT_EG(sr)
        
        .AMPLIT_EG = ReadA_EG12(sr)
        '.OSC1_AMPLIT_EG = ReadA_EG(sr)
        '.OSC2_AMPLIT_EG = ReadA_EG(sr)
        
        .MODULATION_GEN = ReadMG(sr)
        .PORTAMENTO = ReadPORTAMENTO(sr)
        .JOYSTICK = ReadJOYSTICK(sr)
        
        .VELOCITY = ReadVELOCITY(sr)
        .AFTER_TOUCH = ReadAFT_TOUCH(sr)
        
        .ASSIGN_MODE = ReadASSIGNMODE(sr)
        .VOICENAME = ReadVOICENAME(sr)
        
        .MULTIEFFECT = ReadMULTI_EFFECT(sr)
    End With
    b = sr.ReadByte
    'Debug.Print sr.FilePos - 1 - sr.HeaderOffset
End Function
