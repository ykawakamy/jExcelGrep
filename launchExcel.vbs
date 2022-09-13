opentarget = Wscript.Arguments.Item(0)
sheet = Wscript.Arguments.Item(1)
  MsgBox "フルパス: "   & opentarget & sheet

Set xls = WScript.CreateObject("Excel.Application")
xls.visible = True
Set book = xls.Workbooks.Open(opentarget)
xls.Sheets(sheet).Select()

