VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Base64Tab(63) As Byte
Private DecodeTable(233) As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Sub Class_Initialize()

  'initialize the base64 table
  Dim i As Long
  
  Dim tDecodeTable As Variant

    tDecodeTable = Array("255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "62", "255", "255", "255", "63", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "255", "255", "255", "64", "255", "255", "255", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", _
                   "18", "19", "20", "21", "22", "23", "24", "25", "255", "255", "255", "255", "255", "255", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255" _
                   , "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255")
    For i = LBound(tDecodeTable) To UBound(tDecodeTable)
        DecodeTable(i) = tDecodeTable(i)
    Next i
    For i = 65 To 90
        Base64Tab(i - 65) = i
    Next i
    For i = 97 To 122
        Base64Tab(i - 71) = i
    Next i
    For i = 0 To 9
        Base64Tab(i + 52) = 48 + i
    Next i
    Base64Tab(62) = 43
    Base64Tab(63) = 47

End Sub

Public Sub DecodeB64(ByRef FileIn() As Byte, ByRef Out() As Byte)

  'declarations
  
  Dim inp(3) As Byte
  Dim iTemp As Long
  Dim i As Long
  Dim Length As Long
  Dim Remaining As Byte
  Dim BytesOut As Long
  Dim lTemp2 As Long

    If FileIn(UBound(FileIn)) = 61 Then
        Remaining = 1
        If FileIn(UBound(FileIn) - 1) = 61 Then
            Remaining = 2
        End If
    End If
    Length = UBound(FileIn) + 1 'length of the string
    BytesOut = ((Length / 4) * 3) - Remaining ' how many bytes will the decoded string have
    ReDim Out(BytesOut - 1)
    For i = 0 To Length Step 4
        inp(0) = DecodeTable(FileIn(i))
        inp(1) = DecodeTable(FileIn(i + 1))
        inp(2) = DecodeTable(FileIn(i + 2))
        inp(3) = DecodeTable(FileIn(i + 3))
        If inp(3) = 64 Or inp(2) = 64 Then
            If inp(3) = 64 And Not (inp(2) = 64) Then
                inp(0) = DecodeTable(FileIn(i))
                inp(1) = DecodeTable(FileIn(i + 1))
                inp(2) = DecodeTable(FileIn(i + 2))
                '2 bytes out
                Out(iTemp) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
                Out(iTemp + 1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
                Exit Sub
              ElseIf inp(2) = 64 Then
                inp(0) = DecodeTable(FileIn(i))
                inp(1) = DecodeTable(FileIn(i + 1))
                '1 byte out
                Out(iTemp) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
                Exit Sub
            End If
        End If
        '3 bytes out
        Out(iTemp) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
        Out(iTemp + 1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
        Out(iTemp + 2) = ((inp(2) And &H3) * 64) Or inp(3)
        iTemp = iTemp + 3
        'DoEvents
    Next i

End Sub

Public Sub EncodeB64(ByRef FileIn() As Byte, ByRef Out() As Byte)

  'declarations
  
  Dim bin(2) As Byte
  Dim iTemp As Long
  Dim i As Long
  Dim Length As Long
  Dim Remaining As Byte
  Dim BytesOut As Long

    Length = UBound(FileIn) + 1 'length of the string
    Remaining = ((Length) Mod 3)
    If Remaining = 0 Then
        BytesOut = ((Length / 3) * 4)  ' how many bytes will the encoded string have
      Else
        BytesOut = (((Length + (3 - Remaining)) / 3) * 4) ' how many bytes will the encoded string have
    End If
    ReDim Out(BytesOut - 1)
    For i = 0 To Length - Remaining - 1 Step 3
        '3 bytes in
        bin(0) = FileIn(i)
        bin(1) = FileIn(i + 1)
        bin(2) = FileIn(i + 2)
        '4 bytes out
        Out(iTemp) = Base64Tab((bin(0) \ 4) And &H3F)
        Out(iTemp + 1) = Base64Tab((bin(0) And &H3) * 16 Or (bin(1) \ 16) And &HF)
        Out(iTemp + 2) = Base64Tab((bin(1) And &HF) * 4 Or (bin(2) \ 64) And &H3)
        Out(iTemp + 3) = Base64Tab(bin(2) And &H3F)
        iTemp = iTemp + 4
        'DoEvents
    Next i
    If Remaining = 1 Then ' if there is 1 byte remaining
        'read 1 byte, the second in 0
        bin(0) = FileIn(UBound(FileIn))
        bin(1) = 0
        Out(UBound(Out) - 3) = Base64Tab((bin(0) \ 4) And &H3F)
        Out(UBound(Out) - 2) = Base64Tab((bin(0) And &H3) * 16 Or (bin(1) \ 16) And &HF)
        Out(UBound(Out) - 1) = 61
        Out(UBound(Out)) = 61
      ElseIf Remaining = 2 Then 'if there are 2 bytes remaining
        'read 2 bytes, the third is 0
        bin(0) = FileIn(UBound(FileIn) - 1)
        bin(1) = FileIn(UBound(FileIn))
        bin(2) = 0
        Out(UBound(Out) - 3) = Base64Tab((bin(0) \ 4) And &H3F)
        Out(UBound(Out) - 2) = Base64Tab((bin(0) And &H3) * 16 Or (bin(1) \ 16) And &HF)
        Out(UBound(Out) - 1) = Base64Tab((bin(1) And &HF) * 4 Or (bin(2) \ 64) And &H3)
        Out(UBound(Out)) = 61
    End If

End Sub

Public Sub Span(CharsPerLine As Long, InArray() As Byte, OutArray() As Byte)

    Dim liLineCount As Long
    Dim liChunkSize As Long
    Dim liInUbound As Long
    Dim liOutUbound As Long
    Dim liLBound As Long
    Dim liInCounter As Long
    Dim liOutCounter As Long
    Dim liInSize As Long
    Dim liOutSize As Long
    
    liInUbound = UBound(InArray)
    liLBound = LBound(InArray)

    liInSize = liInUbound - liLBound + 1
    

    liLineCount = (liInSize + CharsPerLine - liInSize Mod CharsPerLine) \ CharsPerLine
    
    liOutUbound = liInUbound + (liLineCount * 2)
    liOutSize = liOutUbound - liLBound + 1
    
    ReDim OutArray(liLBound To liOutUbound)
    
    liInCounter = liLBound
    liOutCounter = liLBound
    
    Do While Not liInCounter > liInUbound
        liChunkSize = CharsPerLine
        
        If liChunkSize + liInCounter > liInUbound Then _
           liChunkSize = (liInUbound - liInCounter + 1)
    
        CopyMemory OutArray(liOutCounter), InArray(liInCounter), liChunkSize
        
        liInCounter = liInCounter + liChunkSize
        liOutCounter = liOutCounter + liChunkSize
        
        CopyMemory OutArray(liOutCounter), 2573, 2
        
        liOutCounter = liOutCounter + 2
    
    Loop
End Sub

Public Sub Str2ByteArray(StringIn As String, ByteArray() As Byte)

    ByteArray = StrConv(StringIn, vbFromUnicode)

End Sub

Public Sub Unspan(ArrayIn() As Byte, ArrayOut() As Byte)

  Dim sTemp As String

    sTemp = StrConv(ArrayIn, vbUnicode)
    sTemp = Replace(sTemp, vbCrLf, "")
    ArrayOut = StrConv(sTemp, vbFromUnicode)

End Sub

':) Ulli's VB Code Formatter V2.10.8 (09.03.2002 18:20:28) 2 + 176 = 178 Lines
