On Error Resume Next

Const PAGE_LOADED = 4

dim objIE
dim objshellwindows
set objshellwindows = CreateObject("Shell.Application").Windows

For Each wnd In objshellwindows

                              If InStr(1,wnd.FullName, "iexplore.exe", vbTextCompare) > 0 Then
                              Set objIE = wnd
                              wnd.Navigate2 "[REDACTED_URL]?",2048
                              Do Until wnd.ReadyState = PAGE_LOADED : Call WScript.Sleep(100) : Loop
                              wnd.Document.all.subAcctLoginName.Value = "[REDACTED]"
                              wnd.Document.all.subAcctPassword.Value = "[REDACTED]"
                              Call wnd.Document.all.form1.submit
                              Exit For 
               End If     
Next


objIE.Visible = True

Do Until objIE.ReadyState = PAGE_LOADED : Call WScript.Sleep(100) : Loop

objIE.Document.all.subAcctLoginName.Value = "[REDACTED]"

objIE.Document.all.subAcctPassword.Value = "[REDACTED]"

If Err.Number <> 0 Then

msgbox "Error: " & err.Description

End If

Call objIE.Document.all.form1.submit

Do While objIE.Busy : Call WScript.Sleep(10) : Loop

Dim LinkHref
Dim a
LinkHref = "[REDACTED]"

For Each a In objIE.Document.getElementsByTagName("a")
  If LCase(Right(a.GetAttribute("href"),5)) = LCase(LinkHref) Then
    a.Click
               Exit For  ''# to stop after the first hit
  End If
Next

Set objIE = Nothing