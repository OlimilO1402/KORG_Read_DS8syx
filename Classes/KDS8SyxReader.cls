VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KDS8SyxReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_FileName As String
Private m_FNr      As Integer
Private m_FOffset  As Long
Private m_Skip8thByte As Boolean
Private m_LastSkippedByte As Byte 'oder LastOddByte, oder LastSkippedOddByte
'
Private m_Chunk    As Byte
Private m_ChunkPos As Long

Friend Sub New_(ByVal aFileName As String)
    m_FileName = aFileName
    m_FNr = FreeFile
    Open m_FileName For Binary As m_FNr
End Sub
Private Sub Class_Terminate()
    CClose
End Sub
Public Sub CClose()
    Close m_FNr
    m_FNr = 0
End Sub
Public Function IsEOF() As Boolean
    If m_FNr = 0 Then Exit Function
    IsEOF = EOF(m_FNr)
End Function
Public Property Get FilePos() As Long
    If m_FNr = 0 Then Exit Property
    FilePos = Seek(m_FNr)
End Property
Public Property Get HeaderOffset() As Long
    HeaderOffset = m_FOffset
End Property
Public Property Let HeaderOffset(ByVal value As Long)
    m_FOffset = value
End Property
Public Property Get Is8thByteSkip() As Boolean
    Is8thByteSkip = m_Skip8thByte
End Property
Public Property Let Is8thByteSkip(ByVal value As Boolean)
    m_Skip8thByte = value
End Property
Public Property Get LastSkippedByte() As Byte
    LastSkippedByte = m_LastSkippedByte
End Property

Public Function ReadByte() As Byte
    Dim b As Byte
    Get m_FNr, , b
    ReadByte = b
    
    If m_Skip8thByte Then
        If ((Seek(m_FNr) - m_FOffset) Mod 8) = 0 Then
            Get m_FNr, , b
            m_LastSkippedByte = b
            Debug.Print "&H" & Hex(b)
        End If
    End If
End Function

'Public Function ReadUnsigned8Bit() As Integer
'    ReadUnsigned8Bit = ReadByte
'End Function
Public Function ReadSigned8Bit() As Integer
    ReadSigned8Bit = ReadByte - 127 '128
End Function

'Public Function ReadUnsigned7Bit() As Integer
'    ReadUnsigned8Bit = ReadByte
'End Function
Public Function ReadSigned7Bit() As Integer
    ReadSigned7Bit = ReadByte - 63 '64
End Function

'Public Function ReadUnsigned6Bit() As Integer
'    ReadUnsigned8Bit = ReadByte
'End Function
Public Function ReadSigned6Bit() As Integer
    ReadSigned6Bit = ReadByte - 31 '32
End Function

'Public Function ReadUnsigned5Bit() As Integer
'    ReadUnsigned8Bit = ReadByte
'End Function
Public Function ReadSigned5Bit() As Integer
    ReadSigned5Bit = ReadByte - 15 '32
End Function

Public Function Read7BitChar() As String
    Dim b As Byte: b = ReadByte
    Read7BitChar = Chr$(b + 32)
End Function

Public Function Read7BitString(Optional ByVal strlen As Integer = 10) As String
    Dim s As String
    Dim i As Long
    If strlen > 0 Then
        For i = 0 To strlen - 1
            s = s & Read7BitChar
        Next
    Else
        Dim b As Byte
        Dim c As String
        Const c0 As String = vbNullChar & vbNullChar
        Do While c <> c0
            b = ReadByte
            c = Chr$(b)
            If c = vbNullChar Then
                c = c & Chr$(ReadByte)
            Else
                s = s & Chr$(b + 32)
            End If
        Loop
    End If
    Read7BitString = s
End Function

Public Function SetBack(ByVal sb As Byte) As Long
    Dim b As Byte
    Get m_FNr, FilePos - (sb + 1), b
    SetBack = FilePos
    'seek(m_fnr) = me.FilePos - 1
    'Seek m_FNr, Me.FilePos - 1
End Function

'Private Sub Class_Initialize()
'
'End Sub
'Private Function ReadNextChunk() As Byte
'    If EOF(m_FNr) Then Exit Sub
'    Get m_FNr, , ReadNextChunk
'    m_Chunk = ReadNextChunk
'    m_ChunkPos = 0
'End Function
'
'Private Function ReadNextValue(ByVal size As Integer) As Integer
'    'unsigned
'    m_ChunkPos = m_ChunkPos + size
'    If m_ChunkPos <= 8 Then
'    Select Case size
'    Case 1: ReadNextValue = (m_Chunk And 2 ^ m_ChunkPos)
'    Case 2:
'    Case 4:
'    Case 8:
'    End Select
'    Else
'        'nix
'    End If
'End Function
'
''Reads the next Bit
'Public Function ReadBit() As Integer
'    ReadBit = ReadNextValue(1)
'End Function
'
''We want to have the following functions:
''ReadBit     '1 Bit
''Read2Bit    '2 Bits (unsigned) '
''Read2SBit   '2 Bits signed     '-3, -2, -1, 0, 1, 2
''ReadNibble  '4 Bits (unsigned)
''ReadSNibble '4 Bits signed
''ReadByte    '8 Bits (unsigned)
''ReadSByte   '8 Bits signed
''and we want to use them without the need to take care about the Stream
''chunksize
'
''Reads the next unsigned Byte
'Public Function ReadByte() As Byte
'    If m_ChunkSize = 8 Then
'        ReadByte = ReadNextChunk
'    Else
'
'    End If
'    'm_ChunkSize = 8
'    'ReadByte = m_ChunkByte
'End Function
'
''Reads the next Signed Byte
'Public Function ReadSByte() As Integer
'    Get m_FNr, , m_ChunkByte
'    m_ChunkSize = 8
'    ReadSByte = m_ChunkByte
'End Function
'Public Function ReadNibble() As Byte
'    Get m_FNr, , m_ChunkByte
'    m_ChunkSize = 8
'    ReadSByte = m_ChunkByte
'End Function
'
'
