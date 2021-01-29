Attribute VB_Name = "MClipBoard"
Option Explicit

Public Function KORG_DS8_Prog_ToCBStr(aKORGDS8Prog As KORG_DS8_PROG) As String
    Dim s As String
    With aKORGDS8Prog
        s = s & VOICENAME_ToCBStr(.VOICENAME) '& vbCrLf
        
        s = s & PITCH_ToCBStr(.PITCH) & vbCrLf
        s = s & PITCH_EG_ToCBStr(.PITCH_EG) & vbCrLf
        
        s = s & WFRM1_ToCBStr(.OSC1_WFRM1) & vbCrLf
        s = s & WFRM2_ToCBStr(.OSC2_WFRM2) & vbCrLf
        
        With .TIMBRE_EG
            s = s & T_EG_ToCBStr(.OSC1) & vbCrLf
            s = s & T_EG_ToCBStr(.OSC2) & vbCrLf
        End With
        
        With .AMPLIT_EG
            s = s & A_EG_ToCBStr(.OSC1) & vbCrLf
            s = s & A_EG_ToCBStr(.OSC2) & vbCrLf
        End With
        
        s = s & MG_ToCBStr(.MODULATION_GEN) & vbCrLf
        s = s & PORTAMENTO_ToCBStr(.PORTAMENTO) & vbCrLf
        s = s & JOYSTICK_ToCBStr(.JOYSTICK) & vbCrLf
        
        s = s & VELOCITY_ToCBStr(.VELOCITY) & vbCrLf
        s = s & AFT_TOUCH_ToCBStr(.AFTER_TOUCH) & vbCrLf
        
        s = s & ASSIGNMODE_ToCBStr(.ASSIGN_MODE) & vbCrLf
        s = s & VOICENAME_ToCBStr(.VOICENAME) & vbCrLf
        
        s = s & MULTI_EFFECT_ToCBStr(.MULTIEFFECT) & vbCrLf
    End With
    KORG_DS8_Prog_ToCBStr = s
End Function

Function PITCH_ToCBStr(aPITCH As PITCH) As String
    Dim s As String ': s = vbCrLf
    With aPITCH
        s = s & OSC_ToStr(.OSC1) & vbCrLf    'Oscillator 1 ' 0.5; . . . ; 15;   0.5 (16-Fuss-Orgel); 1 (8-FO); 2 (4-FO); . . .
        s = s & OSC_ToStr(.OSC2) & vbCrLf    'Oscillator 1 ' 0.5; . . . ; 15;   0.5 (16-Fuss-Orgel); 1 (8-FO); 2 (4-FO); . . .
        s = s & SemiNib_ToStr(.DTN) & vbCrLf 'Detune       ' 0; 1; 2; 3
    End With
    PITCH_ToCBStr = s & vbCrLf
End Function

Function PITCH_EG_ToCBStr(aPITCH_EG As PITCH_EG) As String
    Dim s As String ': s = vbCrLf
    With aPITCH_EG
        s = s & Signed_7Bit_ToStr(.STL) & vbCrLf   'Anfangspegel         '-63; . . . ; +63; 'Die Tonhöhe bei der die Klangfarbe beginnt jedesmal wenn eine Note gespielt wird
        s = s & Unsigned_6Bit_ToStr(.ATK) & vbCrLf 'Einschwingungszeit   '  0; . . . ;  63; 'Die Zeit in der sich die Tonhöhe vom Anfangspegel zum einschwingungspegel ändert
        s = s & Signed_7Bit_ToStr(.ATL) & vbCrLf   'Einschwingungspegel  '-63; . . . ; +63; 'Der Spitzenpegelwert der Tonhöhe
        s = s & Unsigned_6Bit_ToStr(.DEC) & vbCrLf 'Ausschwingungszeit   '  0; . . . ;  63; 'Die Zeit in der die Tonhöhe vom Einschwingungspegel zum Normalpegel zurückkehrt, währen die Taste gedrücktgehalten bleibt.
        s = s & Unsigned_6Bit_ToStr(.REL) & vbCrLf 'Abklingung           '  0; . . . ;  63; 'Die Zeit in der die Tonhöht zum Abklingungspegel wechselt, nachdem die Taste losgelassen wurde.
        s = s & Signed_7Bit_ToStr(.RLL) & vbCrLf   'Abklingungspegel     '-63; . . . ; +63; 'Der Pegel zu dem die Tonhöhe wechselt, nachdem die Taste losgelassen wurde.
    End With
    PITCH_EG_ToCBStr = s & vbCrLf
End Function

Function WFRM1_ToCBStr(aWFRM1 As WFRM1) As String
    Dim s As String ': s = vbCrLf
    With aWFRM1
                                                                           '             ' SAW,         RECT,        LiSAW,              LiRECT
        s = s & SemiNib1_ToStr(.TYP) & vbCrLf        'Wellenformart        '1; 2; 3; 4; 1: Sägezahn; 2: Rechteck; 3: Heller Sägezahn; 4: Helles Rechteck;
        s = s & Unsigned_3Bit1_ToStr(.SPCT) & vbCrLf 'Spektrum             '1; . . .; 8;
        s = s & SemiNib_ToStr(.RING) & vbCrLf        'Glocken - Modulation '0; 1; 2; 3;
        s = s & KDS8_ONOFF_ToStr(.LIMT) & vbCrLf     'Begrenzung           'ON; OFF;
        s = s & SemiNib_ToStr(.KBD) & vbCrLf         'Tastatur - Abtastung '0; 1; 2; 3;
    End With
    WFRM1_ToCBStr = s & vbCrLf
End Function

Function WFRM2_ToCBStr(aWFRM2 As WFRM2) As String
    Dim s As String ': s = vbCrLf
    With aWFRM2
        s = s & WFRM2_TYP_ToStr(.TYP) & vbCrLf            'Wellenformart        '1; 2; XMOD;  1: Sägezahn; 2: Rechteck; XMOD: osc2 moduliert osc1;
        s = s & Unsigned_3Bit1_ToStr(.SPCT) & vbCrLf      'Spektrum             '1; . . .; 8;
        s = s & SemiNib_ToStr(.RING) & vbCrLf             'Glocken - Modulation '0; 1; 2; 3;
        s = s & KDS8_ONOFF_ToStr(.LIMT) & vbCrLf          'Begrenzung           'ON; OFF;
        s = s & SemiNib_ToStr(.KBD) & vbCrLf          'Tastatur - Abtastung '0; 1; 2; 3;
    End With
    WFRM2_ToCBStr = s & vbCrLf
End Function

Function T_EG_ToCBStr(aT_EG As T_EG) As String
    Dim s As String ': s = vbCrLf
    With aT_EG
        s = s & Unsigned_7Bit99_ToStr(.TIMB) & vbCrLf 'Klangfarbe                     '0; . . .; 99;
        s = s & Nibble_ToStr(.INT) & vbCrLf           'Intensität                     '0; . . .; 15;
        s = s & Unsigned_5Bit_ToStr(.ATK) & vbCrLf    'ATTACK,   Einschwingungszeit   '0; . . .; 31;
        s = s & Unsigned_5Bit_ToStr(.DEC) & vbCrLf    'DECAY,    Ausschwingungszeit   '0; . . .; 31;
        s = s & Nibble_ToStr(.SUS) & vbCrLf           'SUSTAIN,  Haltepunkt           '0; . . .; 15;
        s = s & Nibble_ToStr(.REL) & vbCrLf           'Release,  Abklingen            '0; . . .; 15;
        s = s & SemiNib_ToStr(.KBD) & vbCrLf          'KEYBOARD, Tastatur - Abtastung '0; 1; 2; 3;
    End With
    T_EG_ToCBStr = s & vbCrLf
End Function

Function A_EG_ToCBStr(aA_EG As A_EG) As String
    Dim s As String ': s = vbCrLf
    With aA_EG
        s = s & Unsigned_6Bit_ToStr(.LEVL) & vbCrLf 'Level, Pegel                   '0; . . .; 63;
        s = s & Unsigned_5Bit_ToStr(.ATK) & vbCrLf  'ATTACK,   Einschwingungszeit   '0; . . .; 31;
        s = s & Unsigned_5Bit_ToStr(.DEC) & vbCrLf  'DECAY,    Ausschwingungszeit   '0; . . .; 31;
        s = s & Nibble_ToStr(.SUS) & vbCrLf         'SUSTAIN,  Haltepegel           '0; . . .; 15;
        s = s & Nibble_ToStr(.REL) & vbCrLf         'Release,  Abklingen            '0; . . .; 15;
        s = s & SemiNib_ToStr(.KBD) & vbCrLf        'KEYBOARD, Tastatur - Abtastung '0; 1; 2; 3;
    End With
    A_EG_ToCBStr = s & vbCrLf
End Function

Function MG_ToCBStr(aMG As MG) As String
    Dim s As String ': s = vbCrLf
    With aMG
        s = s & KDS8_MG_WF_ToStr(.WF) & vbCrLf      'As KDS8_MG_WF    'WAVEFORM, Wellenform              'TRI; SAW; SQUR; S/H
        s = s & Unsigned_6Bit_ToStr(.FREQ) & vbCrLf 'As Byte          'FREQUENCE                         '0; . . .; 63;
        s = s & Unsigned_5Bit_ToStr(.DLY) & vbCrLf  'As Byte          'DELAY                             '0; . . .; 31;
        s = s & Unsigned_6Bit_ToStr(.PTCH) & vbCrLf 'As Byte          'PITCH                             '0; . . .; 63;
        s = s & Unsigned_6Bit_ToStr(.T_A) & vbCrLf  'As Byte          'TIMBRE/AMPLITUDE, Klangfarbe      '0; . . .; 63;
        s = s & KDS8_MG_TASEL_ToStr(.TSEL) & vbCrLf 'As KDS8_MG_TASEL 'TIMBRE SELECT, Klangfarbenwahl    'OFF, 1, 2, 1+2
        s = s & KDS8_MG_TASEL_ToStr(.ASEL) & vbCrLf 'As KDS8_MG_TASEL 'AMPLITUDE SELECT, Amplitudenwahl  'OFF, 1, 2, 1+2
    End With
    MG_ToCBStr = s & vbCrLf
End Function

Function PORTAMENTO_ToCBStr(aPORTAMENTO As PORTAMENTO) As String
    Dim s As String ': s = vbCrLf
    With aPORTAMENTO
        s = s & KDS8_PORTA_MODE_ToStr(.MODE) & vbCrLf 'MODE As KDS8_PORTA_MODE 'Modus '1; 2;
        s = s & Unsigned_6Bit_ToStr(.Time) & vbCrLf   'Time As Byte            'Zeit  '0; . . .; 63;
    End With
    PORTAMENTO_ToCBStr = s & vbCrLf
End Function

Function JOYSTICK_ToCBStr(aJOYSTICK As JOYSTICK) As String
    Dim s As String ': s = vbCrLf
    With aJOYSTICK
        s = s & Signed_5Bit12_ToStr(.BEND_PITCH) & vbCrLf 'As Integer 'Biegung: Tonhöhe            '-12; . . .; +12;
        s = s & SemiNib_ToStr(.BEND_TIMB) & vbCrLf        'As Byte    'Biegung: Klangfarbe         '0; 1; 2; 3;
        s = s & SemiNib_ToStr(.MOD_SPEED) & vbCrLf        'As Byte    'Modulation: Geschwindigkeit '0; 1; 2; 3;
    End With
    JOYSTICK_ToCBStr = s & vbCrLf
End Function

Function VELOCITY_ToCBStr(aVELOCITY As VELOCITY) As String
    Dim s As String ': s = vbCrLf
    With aVELOCITY
        s = s & Unsigned_3Bit_ToStr(.TEG1) & vbCrLf 'As Byte 'Klangfarben EG1 '0; . . .; 7;
        s = s & Unsigned_3Bit_ToStr(.TEG2) & vbCrLf 'As Byte 'Klangfarben EG2 '0; . . .; 7;
        s = s & Unsigned_3Bit_ToStr(.AEG1) & vbCrLf 'As Byte 'Amplituden  EG1 '0; . . .; 7;
        s = s & Unsigned_3Bit_ToStr(.AEG2) & vbCrLf 'As Byte 'Amplituden  EG2 '0; . . .; 7;
    End With
    VELOCITY_ToCBStr = s & vbCrLf
End Function

Function AFT_TOUCH_ToCBStr(aAFT_TOUCH As AFT_TOUCH) As String
    Dim s As String ': s = vbCrLf
    With aAFT_TOUCH
        s = s & Unsigned_3Bit_ToStr(.PMG) & vbCrLf  'As Byte 'PITCH MODULATION GENERATOR, Tonhöhen-Modulationserzeuger '0; . . .; 7;
        s = s & Unsigned_3Bit_ToStr(.TIMB) & vbCrLf 'As Byte 'TIMBRE, Klangfarbe                '0; . . .; 7;
        s = s & Unsigned_3Bit_ToStr(.AMP1) & vbCrLf 'As Byte 'OSC 1-Amplitude                   '0; . . .; 7;
        s = s & Unsigned_3Bit_ToStr(.AMP2) & vbCrLf 'As Byte 'OSC 2-Amplitude                   '0; . . .; 7;
    End With
    AFT_TOUCH_ToCBStr = s & vbCrLf
End Function

Function ASSIGNMODE_ToCBStr(aASSIGNMODE As ASSIGNMODE) As String
    Dim s As String ': s = vbCrLf
    With aASSIGNMODE
        s = s & KDS8_ASSIGNMODE_ToStr(.MODE) & vbCrLf      'As KDS8_ASSIGNMODE  'POLY; UNISON;  '
        If .MODE = KDS8_ASSIGNMODE.kds8_UNISON Then
            s = s & KDS8_UNISON_TRIG_ToStr(.TRIG) & vbCrLf 'As KDS8_UNISON_TRIG 'SINGLE; MULTI; '
            s = s & SemiNib_ToStr(.DETUNE) & vbCrLf        'As Byte             ' 0; 1; 2; 3;   'Die Funktion Detune erlaubt es die Tonhöhen der acht Stimmen, die beim Spielen einer einzelnen Note im Modus UNISON leicht zu verstimmen. Dadurch entsteht ein reicher Chor-Effekt mit einstellbarem Bereich.
        Else
            s = s & vbCrLf
            s = s & vbCrLf
        End If
    End With
    ASSIGNMODE_ToCBStr = s & vbCrLf
End Function

Function VOICENAME_ToCBStr(aVOICENAME As VOICENAME) As String
    Dim s As String ': s = vbCrLf
    With aVOICENAME
        s = s & .name & vbCrLf
    End With
    VOICENAME_ToCBStr = s & vbCrLf
End Function

Function MULTI_EFFECT_ToCBStr(aMULTI_EFFECT As MULTI_EFFECT) As String
    Dim s As String ': s = vbCrLf
    With aMULTI_EFFECT
        s = s & KDS8_EffectType_ToStr(.EffectType) & vbCrLf
        Select Case .EffectType
        Case kds8_MANUAL_DLY
            s = s & MANU_TIME_ToStr(.TIME_MANU) & vbCrLf
            s = s & CStr(.FB) & vbCrLf
            s = s & CStr(.MFRQ) & vbCrLf
            s = s & CStr(.MINT) & vbCrLf
            
            '---
            s = s & CStr(.SPED) & vbCrLf
            s = s & CStr(.DPTH) & vbCrLf
            '---
            
            s = s & CStr(.Level) & vbCrLf
            
        Case kds8_LONG_DLY
            s = s & MANU_TIME_ToStr(.TIME_MANU) & vbCrLf
            s = s & CStr(.FB) & vbCrLf
            
            '---
            s = s & CStr(.MFRQ) & vbCrLf
            s = s & CStr(.MINT) & vbCrLf
            s = s & CStr(.SPED) & vbCrLf
            s = s & CStr(.DPTH) & vbCrLf
            '---
            
            s = s & CStr(.Level) & vbCrLf
        Case kds8_SHORT_DLY
            s = s & MANU_TIME_ToStr(.TIME_MANU) & vbCrLf
            s = s & CStr(.FB) & vbCrLf
            
            '---
            s = s & CStr(.MFRQ) & vbCrLf
            s = s & CStr(.MINT) & vbCrLf
            s = s & CStr(.SPED) & vbCrLf
            s = s & CStr(.DPTH) & vbCrLf
            '---
            
            s = s & CStr(.Level) & vbCrLf
        Case kds8_DOUBLING
            s = s & MANU_TIME_ToStr(.TIME_MANU) & vbCrLf
            
            '----
            s = s & CStr(.FB) & vbCrLf
            s = s & CStr(.MFRQ) & vbCrLf
            s = s & CStr(.MINT) & vbCrLf
            s = s & CStr(.SPED) & vbCrLf
            s = s & CStr(.DPTH) & vbCrLf
            '----
            
            s = s & CStr(.Level) & vbCrLf
        Case kds8_FLANGER
            s = s & MANU_TIME_ToStr(.TIME_MANU) & vbCrLf
            s = s & CStr(.FB) & vbCrLf
            
            '---
            s = s & CStr(.MFRQ) & vbCrLf
            s = s & CStr(.MINT) & vbCrLf
            '---
            
            s = s & CStr(.SPED) & vbCrLf
            s = s & CStr(.DPTH) & vbCrLf
            s = s & CStr(.Level) & vbCrLf
        Case kds8_CHORUS
            s = s & MANU_TIME_ToStr(.TIME_MANU) & vbCrLf
            
            '---
            s = s & CStr(.FB) & vbCrLf
            s = s & CStr(.MFRQ) & vbCrLf
            s = s & CStr(.MINT) & vbCrLf
            '---
            
            s = s & CStr(.SPED) & vbCrLf
            s = s & CStr(.DPTH) & vbCrLf
            s = s & CStr(.Level) & vbCrLf
        End Select
    End With
    MULTI_EFFECT_ToCBStr = s & vbCrLf
End Function


