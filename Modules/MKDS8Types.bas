Attribute VB_Name = "MKDS8Types"
Option Explicit

Public Type PITCH
    OSC1 As Byte '1 'Oscillator 1 '0.5; . . . ; 15;   0.5 (16-Fuss-Orgel); 1 (8-Fuss-Orgel); 2 (4-Fuss-Orgel); . . .
    OSC2 As Byte '1 'Oscillator 2 ' "
    DTN  As Byte '1 'Detune       '0; 1; 2; 3
End Type    'Sum: 3
Public Type PITCH_EG
    STL  As Integer '1 'Anfangspegel         '-63; . . . ; +63; 'Die Tonhöhe bei der die Klangfarbe beginnt jedesmal wenn eine Note gespielt wird
    ATK  As Byte    '1 'Einschwingungszeit   '  0; . . . ;  63; 'Die Zeit in der sich die Tonhöhe vom Anfangspegel zum einschwingungspegel ändert
    ATL  As Integer '1 'Einschwingungspegel  '-63; . . . ; +63; 'Der Spitzenpegelwert der Tonhöhe
    DEC  As Byte    '1 'Ausschwingungszeit   '  0; . . . ;  63; 'Die Zeit in der die Tonhöhe vom Einschwingungspegel zum Normalpegel zurückkehrt, währen die Taste gedrücktgehalten bleibt.
    REL  As Byte    '1 'Abklingung           '  0; . . . ;  63; 'Die Zeit in der die Tonhöht zum Abklingungspegel wechselt, nachdem die Taste losgelassen wurde.
    RLL  As Integer '1 'Abklingungspegel     '-63; . . . ; +63; 'Der Pegel zu dem die Tonhöhe wechselt, nachdem die Taste losgelassen wurde.
End Type       'Sum: 6
Public Enum KDS8_ONOFF
    kds8_OFF
    kds8_ON
End Enum
Public Type WFRM1
    TYP  As Byte '1 'Wellenformart        '1; 2; 3; 4;   1: Sägezahn; 2: Rechteck; 3: Heller Sägezahn; 4: Helles Rechteck;
    SPCT As Byte '1 'Spektrum             '1; . . .; 8;  Ändert die Resonanz des Klanges von einer satten, dunklen Klangfarbe zu einer hellen hohen Klangfarbe
    RING As Byte '1 'Glocken - Modulation '0; 1; 2; 3;   Ein spezieller Effekt, der für die Erzeugung metallischer Klänge herangezogen werden kann; ideal für Glocken- oder Becken-Klangfarben
    LIMT As KDS8_ONOFF '1 'Begrenzung     'ON; OFF;      Begrenzt den Wert der Klangfarben-Modulation
    KBD  As Byte '1 'Tastatur - Abtastung '0; 1; 2; 3;   Stellt den Wert ein, mit dem sich der klang von OSC1 im gesamten Tastaturbereich ändert. Der Klang hellt sich auf wenn höhere Noten gespielt wreden und wird weicher sobald tiefere Noten gespielt werden.
End Type    'Sum: 5

Public Type WFRM2
    TYP  As Byte '1 'Wellenformart        '1; 2; XMOD;   1: Sägezahn; 2: Reckteck; XMOD: OSC1 moduliert OSC2 um komplizierte Wellenformen zu produzieren
    SPCT As Byte '1 'Spektrum             '1; . . .; 8;  Ändert die Resonanz des Klanges von einer satten, dunklen Klangfarbe zu einer hellen hohen Klangfarbe
    RING As Byte '1 'Glocken - Modulation '0; 1; 2; 3;   Ein spezieller Effekt, der für die Erzeugung metallischer Klänge herangezogen werden kann; ideal für Glocken- oder Becken-Klangfarben
    LIMT As KDS8_ONOFF '1 'Begrenzung     'ON; OFF;      Begrenzt den Wert der Klangfarben-Modulation
    KBD  As Byte '1 'Tastatur - Abtastung '0; 1; 2; 3;   Stellt den Wert ein, mit dem sich der klang von OSC1 im gesamten Tastaturbereich ändert. Der Klang hellt sich auf wenn höhere Noten gespielt wreden und wird weicher sobald tiefere Noten gespielt werden.
End Type    'Sum: 5

'Public Type EG
'    ATK  As Byte '1 'ATTACK
'    DEC  As Byte '1 'DECAY
'    SUS  As Byte '1 'SUSTAIN
'    REL  As Byte '1 'RELEASE
'    KBD  As Byte '1 'KEYBOARD
'End Type
Public Type T_EG    'TIMBRE EG Schwingungsklangfarben-EG
    TIMB As Byte '1 'Klangfarbe                     '0; . . .; 99;
    INT  As Byte '1 'Intensität                     '0; . . .; 15;
    'MyEG As EG
    ATK  As Byte '1 'ATTACK,   Einschwingungszeit   '0; . . .; 31;
    DEC  As Byte '1 'DECAY,    Ausschwingungszeit   '0; . . .; 31;
    SUS  As Byte '1 'SUSTAIN,  Haltepunkt           '0; . . .; 15;
    REL  As Byte '1 'Release,  Abklingen            '0; . . .; 15;
    KBD  As Byte '1 'KEYBOARD, Tastatur - Abtastung '0; 1; 2; 3;
End Type    'Sum: 7
Public Type T_EG12
    OSC1 As T_EG
    OSC2 As T_EG
End Type

Public Type A_EG 'AMPLITUDE EG Schwingungsamplituden-EG
    LEVL As Byte '1 'Level, Pegel                   '0; . . .; 63;
    'MyEG As EG
    ATK  As Byte '1 'ATTACK,   Einschwingungszeit   '0; . . .; 31;
    DEC  As Byte '1 'DECAY,    Ausschwingungszeit   '0; . . .; 31;
    SUS  As Byte '1 'SUSTAIN,  Haltepegel           '0; . . .; 15;
    REL  As Byte '1 'Release,  Abklingen            '0; . . .; 15;
    KBD  As Byte '1 'KEYBOARD, Tastatur - Abtastung '0; 1; 2; 3;
End Type    'Sum: 6
Public Type A_EG12
    OSC1 As A_EG
    OSC2 As A_EG
End Type

Public Enum KDS8_MG_WF
    kds8_MG_TRI  'TRI, Triangle
    kds8_MG_SAW  'SAW, Sawtooth
    kds8_MG_SQUR 'SQUR, Square
    kds8_MG_S_H  'S/H, Sample&Hold
End Enum
Public Enum KDS8_MG_TASEL
    MGTA_OFF
    MGTA_1
    MGTA_2
    MGTA_1_2
End Enum
Public Type MG      'MODULATION GENERATOR Modulationserzeuger
    WF   As KDS8_MG_WF    'WAVEFORM, Wellenform              'TRI; SAW; SQUR; S/H
    FREQ As Byte          'FREQUENCE                         '0; . . .; 63;
    DLY  As Byte          'DELAY                             '0; . . .; 31;
    PTCH As Byte          'PITCH                             '0; . . .; 63;
    T_A  As Byte          'TIMBRE/AMPLITUDE, Klangfarbe      '0; . . .; 63;
    TSEL As KDS8_MG_TASEL 'TIMBRE SELECT, Klangfarbenwahl    'OFF, 1, 2, 1+2
    ASEL As KDS8_MG_TASEL 'AMPLITUDE SELECT, Amplitudenwahl  'OFF, 1, 2, 1+2
End Type
Public Enum KDS8_PORTA_MODE
    kds8_PortaMode_1
    kds8_PortaMode_2
End Enum
Public Type PORTAMENTO
    MODE As KDS8_PORTA_MODE 'Modus                             '1; 2;
    Time As Byte            'Zeit                              '0; . . .; 63;
End Type
Public Type JOYSTICK
    BEND_PITCH As Integer 'Biegung: Tonhöhe            '-12; . . .; +12;
    BEND_TIMB  As Byte    'Biegung: Klangfarbe         '0; 1; 2; 3;
    MOD_SPEED  As Byte    'Modulation: Geschwindigkeit '0; 1; 2; 3;
End Type
Public Type VELOCITY
    TEG1 As Byte 'Klangfarben EG1                   '0; . . .; 7;
    TEG2 As Byte 'Klangfarben EG2                   '0; . . .; 7;
    AEG1 As Byte 'Amplituden EG1                    '0; . . .; 7;
    AEG2 As Byte 'Amplituden EG2                    '0; . . .; 7;
End Type
Public Type AFT_TOUCH
    PMG  As Byte 'PITCH MODULATION GENERATOR, Tonhöhen-Modulationserzeuger '0; . . .; 7;
    TIMB As Byte 'TIMBRE, Klangfarbe                '0; . . .; 7;
    AMP1 As Byte 'OSC 1-Amplitude                   '0; . . .; 7;
    AMP2 As Byte 'OSC 2-Amplitude                   '0; . . .; 7;
End Type
Public Enum KDS8_ASSIGNMODE
    kds8_POLY
    kds8_UNISON
End Enum
Public Enum KDS8_UNISON_TRIG
    kds8_SINGLE
    kds8_MULTI
End Enum
Public Type ASSIGNMODE
    MODE   As KDS8_ASSIGNMODE  'POLY; UNISON;  '
    TRIG   As KDS8_UNISON_TRIG 'SINGLE; MULTI; '
    DETUNE As Byte             ' 0; 1; 2; 3;   'Die Funktion Detune erlaubt es die Tonhöhen der acht Stimmen, die beim Spielen einer einzelnen Note im Modus UNISON leicht zu verstimmen. Dadurch entsteht ein reicher Chor-Effekt mit einstellbarem Bereich.
End Type
Public Type VOICENAME
    name As String       '10
End Type

'Public Type EFFECT_PARAMS
'    Byt1 As Byte
'    Byt2 As Byte
'    Byt3 As Byte
'    Byt4 As Byte
'    Byt5 As Byte
'End Type
'EffectTypes:
'MANUAL DELAY
'LONG DELAY
'SHORT DELAY
'DOUBLING
'FLANGER
'CHORUS
'Public Type MANUAL_DELAY
'    Time  As Byte    '1 'Verzögerungszeit    0.04; . . .; 850; ms
'    FB    As Integer '1 'FEEDBACK Rückkopplung   -15; . . . ; +15;
'    MFRQ  As Byte    '1 'MODULATION FREQUENCE Modulationsfrequenz    0; . . .; 31;
'    MINT  As Byte    '1 'MODULATION INTENSITY Modulationsintensität  0; . . .; 15;
'    Level As Byte    '1 'Pegel   0; . . .; 31;
'End Type        'Sum: 5
'Public Type LONG_DELAY
'    Time  As Byte    '1 'Verzögerungszeit    105; . . . ; 720; ms
'    FB    As Integer '1 'FEEDBACK Rückkopplung   -15; . . .; +15;
'    Level As Byte    '1 'Pegel   0; . . .; 31;
'End Type        'Sum: 3
'Public Type SHORT_DELAY
'    Time  As Byte    '1 'Verzögerungszeit    20; . . .; 88; ms
'    FB    As Integer '1 'FEEDBACK Rückkopplung   -15; . . .; +15;
'    Level As Byte    '1 'Pegel   0; . . .; 31;
'End Type        'Sum: 3
'Public Type DOUBLING
'    Time  As Byte    '1 'Verzögerungszeit
'    Level As Byte    '1 'Pegel
'End Type        'Sum: 2
'Public Type FLANGER
'    MANU  As Byte    '1 'MANUAL Manuell  0.04; . . .; 5.5; ms    Bestimmt die verzögerungszeit zwischen dem Direkt- und dem Flanger-Signal.
'    FB    As Integer '1 'FEEDBACK Rückkopplung   -15; . . .; +15;    Bestimmt den Wert, mit dem das Flanger-Signal mit sich selbst zurückgekoppelt wird und bestimmt die Intensität des Effekts. Negative Einstellungen produzieren ein phasen-umgekehrtes Flanging, was in einen klareren, helleren Sound resultieren kann.
'    SPED  As Byte    '1 'SPEED Flanger-Modulationsgeschwindigkeit    0; . . .; 24;   Bestimmt die Intensität des Flanger-Effektes, der von einer langsamen, ruhigen Modulation bis zu einer schnellen, extremen Einstellung reicht.
'    DPTH  As Byte    '1 'DEPTH Flanger-Modulationstiefe  0; . . .; 31;   Bestimmt die Flanger-Modulationstiefe, die von Null (kein Effekt) bi zu einem hochmodulierten Sound reicht.
'    Level As Byte    '1 'Pegel   0; . . .; 15;   Bestimmt den Gesamtpegel des Flanger-Signals.
'End Type        'Sum: 5
'Public Type CHORUS
'    MANU  As Byte    '1 'MANUAL  5.0; . . .; 32; ms  Bestimmt die Verzögerungszeit zwischen dem Direkt- und dem Chorus-Signal
'    SPED  As Byte    '1 'SPEED   0; . . .; 31;   Bestimmt die Intensität des Chorus-Effektes, der von einer langsamen, ruhigen Modulation bis zu einer schnellen, extremen Einstellung reicht.
'    DPTH  As Byte    '1 'DEPTH   0; . . .; 31;   Bestimmt die Chorus-Modulationstiefe, die von Null (kein Effekt) bis zu einem hochmodulierten Sound reicht.
'    Level As Byte    '1 'Pegel   0; . . .; 31;   Bestimmt den Gesamtpegl des Chorus-Signals.
'End Type        'Sum: 4
Public Enum KDS8_EffectType
    kds8_MANUAL_DLY
    kds8_LONG_DLY
    kds8_SHORT_DLY
    kds8_DOUBLING
    kds8_FLANGER
    kds8_CHORUS
End Enum
Public Type MULTI_EFFECT
    EffectType As KDS8_EffectType
    TIME_MANU  As Byte
    FB         As Integer
    MFRQ       As Byte
    MINT       As Byte
    SPED       As Byte
    DPTH       As Byte
    Level      As Byte
End Type


Public Type KORG_DS8_PROG
    
    PITCH          As PITCH        '3
    PITCH_EG       As PITCH_EG     '6  '9
    
    OSC1_WFRM1      As WFRM1       '4  '13
    OSC2_WFRM2      As WFRM2       '4  '18
    
    TIMBRE_EG       As T_EG12
    'OSC1_TIMBRE_EG As T_EG         '7  '24
    'OSC2_TIMBRE_EG As T_EG         '7  '31
    
    AMPLIT_EG       As A_EG12
    'OSC1_AMPLIT_EG As A_EG         '6  '37
    'OSC2_AMPLIT_EG As A_EG         '6  '43
    
    MODULATION_GEN As MG           '7  '50
    PORTAMENTO     As PORTAMENTO   '2  '52
    JOYSTICK       As JOYSTICK     '3  '55
    
    VELOCITY       As VELOCITY     '4  '59
    AFTER_TOUCH    As AFT_TOUCH    '4  '63
    
    ASSIGN_MODE    As ASSIGNMODE   '3  '66
    VOICENAME      As VOICENAME    '10 '76
    
    MULTIEFFECT    As MULTI_EFFECT '8  '84
    ' es dürften nur 80bytes sein
    'laut KORG DS-8 Handbuch sollten es 96 sein
    'es dürfen aber definitiv nur 80 sein,
    '==>> d.h. irgendwo steckt ein Datenproblem
End Type

Public Type ListOf_KORG_DS8_PROG
    Count As Long
    Arr() As KORG_DS8_PROG
End Type

Public Type KORG_DS8_KOMBI
    '
    Dummy_Value As Long
End Type

Public Type ListOf_KORG_DS8_KOMB
    Count As Long
    Arr() As KORG_DS8_KOMBI
End Type

