function urldecode(Source)

    Dim obj
    Dim strDecode
    Dim strOutput

    On Error Resume Next
    sTmp=""
    iCount = 1
    lSrcLen=Len(Source)
    Do Until iCount > lSrcLen
        sChr = Mid(Source,iCount,1)
        iCount = iCount+1
        If sChr="+" Then
            sChr = " "
        ElseIf sChr="%" Then
            sHex = Mid(Source,iCount,2)
            iCount = iCount + 2
            iAsc = CByte("&H" & sHex)
            If (&H00 <= iAsc And iAsc <= &H80) Or _
               (&HA0 <= iAsc And iAsc <= &HDF) Then
                '1バイト文字
                sChr=Chr(iAsc)
            ElseIf (&H81 <= iAsc And iAsc <= &H9F) Or _
               (&HE0 <= iAsc And iAsc <= &HFF) Then
                '2バイト文字
                sChr = Mid(Source,iCount,1)
               iCount = iCount + 1
                If sChr="%" Then
                    sHex2 = Mid(Source,iCount,2)
                    iCount = iCount + 2
                Else
                    sHex2 = Hex(Asc(sChr))
                    If Len(sHex2) = 1 Then
                        sHex2 = "0" & sHex2
                    End If
                End If
                sChr=Chr(CInt("&H" & sHex & sHex2))
            End If
        End If
        sTmp=sTmp & sChr
    Loop
    urldecode = sTmp
End function

On Error Resume Next

opentarget = urldecode(Wscript.Arguments.Item(0))
sheetName = urldecode(Wscript.Arguments.Item(1))
cellAddress = urldecode(Wscript.Arguments.Item(2))

'MsgBox "フルパス: "   & opentarget &  "[" & sheetName & "]" 

set xls = GetObject(,"Excel.Application")
If TypeName(xls) = "Empty" Then
  Set xls = WScript.CreateObject("Excel.Application")
End If
On Error GoTo 0
xls.visible = True
Set book = xls.Workbooks.Open(opentarget)

set sheet = xls.Sheets(sheetName)
sheet.Select()

'If TypeName(cellAddress) = "Empty" Then
  sheet.range(cellAddress).Select()
'end if
