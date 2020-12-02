Attribute VB_Name = "MNibble"
Option Explicit

Function LoNib(ByVal b As Byte) As Byte
    'delivers the lower 4 of 8 bits
    LoNib = b And &HF
End Function
Function HiNib(ByVal b As Byte) As Byte
    'delivers the higher 4 of 8 bits
    HiNib = (b \ &H10) And &HF '0 ') / &HF
End Function

Public Function LoSemiNib(ByVal nib As Byte) As Byte
    'delivers the lower 2 of 4 bits
    LoSemiNib = nib And &H3
End Function
Public Function HiSemiNib(ByVal nib As Byte) As Byte
    'delivers the higher 2 of 4 bits
    HiSemiNib = (nib \ &H4) And &H3
End Function

Public Function Lo3Bit6(ByVal b As Byte) As Byte
    'delivers the lower 3 of 6 bits
    Lo3Bit6 = b And &H7
End Function
Public Function Hi3Bit6(ByVal b As Byte) As Byte
    'delivers the higher 3 of 6 bits
    Hi3Bit6 = (b \ &H8) And &H7
End Function

Public Function Lo2Bit8(ByVal b As Byte) As Byte
    'delivers the lower 2 of 8 bits
    Lo2Bit8 = b And &H3
End Function
Public Function Hi6Bit8(ByVal b As Byte) As Byte
    'delivers the higher 6 of 8 bits
    Hi6Bit8 = (b \ &H4) And &H3F
End Function

Public Function Lo3Bit8(ByVal b As Byte) As Byte
    'delivers the lower 3 of 6 bits
    Lo3Bit8 = b And &H7
End Function
Public Function Hi5Bit8(ByVal b As Byte) As Byte
    'delivers the higher 5 of 8 bits
    Hi5Bit8 = (b \ &H8) And &H1F
End Function

Public Function Lo5Bit8(ByVal b As Byte) As Byte
    'delivers the lower 5 of 8 bits
    Lo5Bit8 = b And &H1F
End Function
Public Function Hi3Bit8(ByVal b As Byte) As Byte
    'delivers the higher 3 of 8 bits
    Hi3Bit8 = (b \ &H20) And &H7
End Function

Public Function Lo6Bit8(ByVal b As Byte) As Byte
    'delivers the lower 6 of 8 bits
    Lo6Bit8 = b And &H3F
End Function
Public Function Hi2Bit8(ByVal b As Byte) As Byte
    'delivers the higher 2 of 8 bits
    Hi2Bit8 = (b \ &H40) And &H3
End Function

Function SemiNib_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 0 To 3: s = CStr(value)
    Case Else:   s = "&H" & Hex$(value)
    End Select
    SemiNib_ToStr = s
End Function
Function SemiNib1_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 1 To 4: s = CStr(value)
    Case Else:   s = "&H" & Hex$(value)
    End Select
    SemiNib1_ToStr = s
End Function




Function Signed_3Bit_ToStr(ByVal value As Integer) As String
    Dim s As String
    Select Case value
    Case -3 To 3: s = CStr(value)
    Case Else:   s = "&H" & Hex$(value)
    End Select
    Signed_3Bit_ToStr = s
End Function
Function Unsigned_3Bit_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 0 To 7: s = CStr(value)
    Case Else:   s = "&H" & Hex$(value)
    End Select
    Unsigned_3Bit_ToStr = s
End Function
Function Unsigned_3Bit1_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 1 To 8: s = CStr(value)
    Case Else:   s = "&H" & Hex$(value)
    End Select
    Unsigned_3Bit1_ToStr = s
End Function


Function Nibble_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 0 To 15: s = CStr(value)
    Case Else:    s = "&H" & Hex$(value)
    End Select
    Nibble_ToStr = s
End Function
Function Signed_5Bit_ToStr(ByVal value As Integer) As String
    Dim s As String
    Select Case value
    Case -15 To 15: s = CStr(value)
    Case Else:      s = "&H" & Hex$(value)
    End Select
    Signed_5Bit_ToStr = s
End Function
Function Signed_5Bit12_ToStr(ByVal value As Integer) As String
    Dim s As String
    Select Case value
    Case -12 To 12: s = CStr(value)
    Case Else:      s = "&H" & Hex$(value)
    End Select
    Signed_5Bit12_ToStr = s
End Function

Function Unsigned_5Bit_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 0 To 31: s = CStr(value)
    Case Else:    s = "&H" & Hex$(value)
    End Select
    Unsigned_5Bit_ToStr = s
End Function
Function Signed_6Bit_ToStr(ByVal value As Integer) As String
    Dim s As String
    Select Case value
    Case -31 To 31: s = CStr(value)
    Case Else:      s = "&H" & Hex$(value)
    End Select
    Signed_6Bit_ToStr = s
End Function
Function Unsigned_6Bit_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 0 To 63: s = CStr(value)
    Case Else:    s = "&H" & Hex$(value)
    End Select
    Unsigned_6Bit_ToStr = s
End Function
Function Signed_7Bit_ToStr(ByVal value As Integer) As String
    Dim s As String
    Select Case value
    Case -63 To 63: s = CStr(value)
    Case Else:      s = "&H" & Hex$(value)
    End Select
    Signed_7Bit_ToStr = s
End Function
Function Unsigned_7Bit_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 0 To 127: s = CStr(value)
    Case Else:     s = "&H" & Hex$(value)
    End Select
    Unsigned_7Bit_ToStr = s
End Function

Function Unsigned_7Bit99_ToStr(ByVal value As Byte) As String
    Dim s As String
    Select Case value
    Case 0 To 99: s = CStr(value)
    Case Else:    s = "&H" & Hex$(value)
    End Select
    Unsigned_7Bit99_ToStr = s
End Function

