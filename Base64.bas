Attribute VB_Name = "Base64"
'BASE64编码模块

Option Explicit

'外部直接调用此函数即可
Public Sub Base64Encode(ByRef pbAInput() As Byte, ByRef sOut As String, Optional ByRef sPrexSpace As String = "", Optional ByRef charsPerLine As Integer = 68)
    
    Dim abOut() As Byte, idx As Long, sBase64 As String
    Dim lengthMod As Integer
    Dim paddingCount As Integer
    
    Encode pbAInput, abOut
    
    ByteArrayToString abOut, sBase64
    sOut = ""
    If Len(sBase64) = 0 Then Exit Sub
    
    'Base64要求长度为4的倍数，如果不够，则添加"="
    lengthMod = Len(sBase64) Mod 4
    paddingCount = 4 - lengthMod
    sBase64 = sBase64 & String(paddingCount, "=")
    
    '转换为合适的换行模式
    If charsPerLine > 0 Then
        For idx = 1 To Len(sBase64) Step charsPerLine
            sOut = sOut & sPrexSpace & Mid(sBase64, idx, charsPerLine) & vbCrLf
        Next
        sOut = Mid(sOut, 1, Len(sOut) - 3) '去掉上面语句添加的最后的回车换行
    Else
        sOut = sBase64
    End If
    
End Sub

Public Sub ByteArrayToString(ByRef pbArrayInput() As Byte, ByRef strOut As String)
    strOut = StrConv(pbArrayInput, vbUnicode)
End Sub

Public Sub Encode(ByRef pbArrayInput() As Byte, ByRef pbArrayOutput() As Byte)
    Dim iSizeMod As Integer, lSizeIn  As Long, lSizeOut As Long, Index    As Long
    Dim lIndex2  As Long, lTotal   As Long, bBuffer(2) As Byte
    Dim mvB64Enc(63) As Byte
    Dim iIndex As Integer
    
    For iIndex = 65 To 90
        mvB64Enc(iIndex - 65) = iIndex
    Next
    For iIndex = 97 To 122
        mvB64Enc(iIndex - 71) = iIndex
    Next
    For iIndex = 48 To 57
        mvB64Enc(iIndex + 4) = iIndex
    Next
    mvB64Enc(62) = 43
    mvB64Enc(63) = 47
    
    lSizeIn = UBound(pbArrayInput) + 1
    iSizeMod = lSizeIn Mod 3
    lSizeOut = ((lSizeIn - iSizeMod) \ 3) * 4
    If iSizeMod > 0 Then lSizeOut = lSizeOut + 4
    
    ReDim pbArrayOutput(lSizeOut - 1)
    
    If lSizeIn >= 3 Then
        
        lTotal = lSizeIn - iSizeMod - 1
        For Index = 0 To lTotal Step 3
            
            bBuffer(0) = pbArrayInput(Index)
            bBuffer(1) = pbArrayInput(Index + 1)
            bBuffer(2) = pbArrayInput(Index + 2)
            pbArrayOutput(lIndex2) = mvB64Enc((bBuffer(0) And &HFC) \ 4)
            pbArrayOutput(lIndex2 + 1) = mvB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
            pbArrayOutput(lIndex2 + 2) = mvB64Enc((bBuffer(1) And &HF) * 4 Or (bBuffer(2) And &HC0) \ 64)
            pbArrayOutput(lIndex2 + 3) = mvB64Enc((bBuffer(2) And &H3F))
            lIndex2 = lIndex2 + 4
        Next
    End If
    
    Select Case iSizeMod
    Case 1
        bBuffer(0) = pbArrayInput(lSizeIn - 1)
        bBuffer(1) = 0
        pbArrayOutput(lIndex2) = mvB64Enc((bBuffer(0) And &HFC) \ 4)
        pbArrayOutput(lIndex2 + 1) = mvB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
        pbArrayOutput(lIndex2 + 2) = 61
        pbArrayOutput(lIndex2 + 3) = 61
    Case 2
        bBuffer(0) = pbArrayInput(lSizeIn - 2)
        bBuffer(1) = pbArrayInput(lSizeIn - 1)
        bBuffer(2) = 0
        pbArrayOutput(lIndex2) = mvB64Enc((bBuffer(0) And &HFC) \ 4)
        pbArrayOutput(lIndex2 + 1) = mvB64Enc((bBuffer(0) And &H3) * 16 Or (bBuffer(1) And &HF0) \ 16)
        pbArrayOutput(lIndex2 + 2) = mvB64Enc((bBuffer(1) And &HF) * 4 Or (bBuffer(2) And &HC0) \ 64)
        pbArrayOutput(lIndex2 + 3) = 61
    End Select
End Sub

