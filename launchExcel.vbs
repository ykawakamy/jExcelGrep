function urldecode(Source)

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
                sChr=Chr(iAsc)
            ElseIf (&H81 <= iAsc And iAsc <= &H9F) Or _
               (&HE0 <= iAsc And iAsc <= &HFF) Then
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

function urldecodeUTF8(Source)
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
            If (&H00 <= iAsc And iAsc <= &H7F) Then
                sChr=Chr(iAsc)
            ELSEIf (&HC2 <= iAsc And iAsc <= &HDF) Then
              iCount = iCount + 1
              sHex = Mid(Source,iCount,2)
              iCount = iCount + 2
              iAsc = &H80 + (iAsc - &HC2) * 64 + (CByte("&H" & sHex) - &H80 )
              sChr=ChrW(iAsc)
'              msgbox iAsc & ";1 " & sChr
            ELSEIf (&HE0 <= iAsc And iAsc <= &HEF) Then
              iCount = iCount + 1
              sHex = Mid(Source,iCount,2)
              iCount = iCount + 2
              iCount = iCount + 1
              sHex2 = Mid(Source,iCount,2)
              iCount = iCount + 2
              iAsc = (iAsc - &HE0) * 64 * 64 + (CByte("&H" & sHex) - &H80 ) * 64 + (CByte("&H" & sHex2) - &H80 )
              sChr=ChrW(iAsc)
'              msgbox iAsc & ";2 " & sChr & sHex & sHex2
            End If
        End If
        sTmp=sTmp & sChr
    Loop
    urldecodeUTF8 = sTmp
End function

function checkDefined(value)
 On Error Resume Next
  checkDefined = True
  typeNameStr = TypeName(value)
  if typeNameStr = "Empty" then 
    checkDefined = False
  elseif typeNameStr = "Nothing" then
    checkDefined = False
  elseif typeNameStr = "String" then
    if value = "" then 
      checkDefined = False
    end if
  end if
'  msgbox typeNameStr & checkDefined
end function 

On Error Resume Next

opentarget = (Wscript.Arguments.Item(0))
filename = (Wscript.Arguments.Item(1))
sheetName = (Wscript.Arguments.Item(2))
cellAddress = (Wscript.Arguments.Item(3))
' msgbox opentarget & " : " & filename & " : " & sheetName & " : " & cellAddress

opentarget = urldecodeUTF8(opentarget)
filename = urldecodeUTF8(filename)
sheetName = urldecodeUTF8(sheetName)
cellAddress = urldecodeUTF8(cellAddress)

'msgbox opentarget & " : " & filename & " : " & sheetName & " : " & cellAddress
set xls = GetObject(,"Excel.Application")
If TypeName(xls) = "Empty" Then
  Set xls = WScript.CreateObject("Excel.Application")
End If

xls.visible = True
'On Error Resume Next

Set book = xls.Workbooks(filename)
if checkDefined(book) Then
  if book.path <> opentarget Then
    msgbox "すでに同じ名のファイルがオープンされています。"
    WScript.Quit
  end if

end if

On Error GoTo 0
If not checkDefined(book) Then
  Set book = xls.Workbooks.Open(opentarget & "\" & filename)
  If not checkDefined(book) Then
    msgbox "ファイルオープンに失敗しました。"
    WScript.Quit
  end if
end if


book.activate

If checkDefined(sheetName) Then
  'msgbox "[" & sheetName & "]"
  set sheet = xls.Sheets(sheetName)
  sheet.Activate()
  If checkDefined(cellAddress) Then
    sheet.range(cellAddress).Activate()
  end if
end if
