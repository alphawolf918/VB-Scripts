Dim wshShell
Dim ieObject
Dim LinkHref
Dim a
Dim NSN

Set wshShell = WScript.CreateObject("WScript.Shell")
Set ieObject = CreateObject("InternetExplorer.Application")
LinkHref = "HAYXF"

Const PAGE_LOADED = 4

ieObject.Visible = True
ieObject.Navigate "[REDACTED_URL]?"
Do Until ieObject.ReadyState = PAGE_LOADED : Call WScript.Sleep(100) : Loop
ieObject.Document.All.Item("subAcctLoginName").Value = "[REDACTED]"
ieObject.Document.All.Item("subAcctPassword").Value = "[REDACTED]"
ieObject.Document.All.Item("form1").Submit

Do While ieObject.Busy : Call WScript.Sleep(10) : Loop
For Each a In ieObject.Document.getElementsByTagName("a")
  If LCase(Right(a.GetAttribute("href"),5)) = LCase(LinkHref) Then
    a.Click
	Exit For
  End If
Next

NSN = InputBox("Please enter an NSN:","NSN Entry")
Do While ieObject.Busy : Call WScript.Sleep(5) : Loop
	If NSN <> "" And NSN <> "undefined" Then
		ieObject.Document.All.Item("FNiin").Value = NSN
		ieObject.Document.All.Item("FrmFlis").Submit
	Else
		MsgBox "No NSN was entered. Form will not submit."
	End If

Set ieObject = Nothing
Set wshShell = Nothing
Set NSN = Nothing