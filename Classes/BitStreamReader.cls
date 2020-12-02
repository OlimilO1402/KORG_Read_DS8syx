VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitStreamReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_byteArray() As Byte
Private m_bufferLengthInBits As Long
Private m_byteArrayIndex As Long
Private m_partialByte As Byte
Private m_cbitsInPartialByte As Byte


Public Property Get EndOfStream() As Boolean
    EndOfStream = 0 = m_bufferLengthInBits
End Property
Public Property Get CurrentIndex() As Long
    CurrentIndex = m_byteArrayIndex - 1
End Property
Public Property Get ReadByte(ByVal countOfBits As Byte) As Byte
    If (Me.EndOfStream) Then
        MsgBox "EndOfStreamReached"
    End If
    If (countOfBits <= 0 Or 8 < countOfBits) Then
        MsgBox "countOfBits"
    End If
    If (countOfBits > m_bufferLengthInBits) Then
        MsgBox "countOfBits"
    End If
    m_bufferLengthInBits = m_bufferLengthInBits - countOfBits
    Dim b As Byte
    If (m_cbitsInPartialByte >= countOfBits) Then
        Dim num As Long: num = 8 - countOfBits
        b = ShR(m_partialByte, num) ' >>
        m_partialByte = ShL(m_partialByte, countOfBits) '<<
        m_cbitsInPartialByte = m_cbitsInPartialByte - countOfBits
    Else
        Dim b2 As Byte: b2 = m_byteArray(m_byteArrayIndex)
        m_byteArrayIndex = m_byteArrayIndex + 1
        Dim num2 As Long: num2 = 8 - countOfBits
        b = ShR(m_partialByte, num2) '>>
        Dim num3 As Long: num3 = Math.Abs(countOfBits - m_cbitsInPartialByte - 8)
        b = b Or ShR(b2, num3) '>>
        m_partialByte = ShR(b2, countOfBits - m_cbitsInPartialByte) ' <<
        m_cbitsInPartialByte = 8 - (countOfBits - m_cbitsInPartialByte)
    End If
    ReadByte = b
End Property

