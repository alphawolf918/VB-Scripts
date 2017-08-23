Sub Extract_NSN_List()

    Dim lrow As Integer
    Dim xlApp As Object, xlWkbk As Object, xlWksht As Object
Set xlApp = GetObject(, "Excel.Application")
Dim XReg As RegExp
    Dim xMatches As MatchCollection
    Dim xMatch As Match

Set XReg = New RegExp
With XReg
        .IgnoreCase = True
        .Global = True
    End With

Set xlWkbk = xlApp.Workbooks.Open("[REDACTED]")
xlWkbk.Activate

Set xlWksht = xlWkbk.Worksheets("Sheet1")
'Set objReg = CreateObject(“vbscript.regexp”)

lrow = 1479
    Dim xlRange As Excel.Range

Set xlRange = xlWksht.Range("C2:C" & lrow)
Dim NSN As String  'Lookup string
    Dim strNSNInfo As String   'Return string
    Dim cRow As Integer   'Current row of spreadsheet
    Dim strSel As String
    cRow = 0
    '(B) Loop through NSN's in worksheet, find in Word Document, add bookmark at beginning of line, add hyperlink to bookmark in Excel.
    For Each Cell In xlRange
        Selection.HomeKey Unit:=wdStory
NSN = Cell.Value
        Debug.Print "Copying NSN " & NSN & "from cell " & Cell.Address
cRow = Cell.Row
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = NSN
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        Selection.Find.Execute

        With Selection
            .HomeKey Unit:=wdLine, Extend:=wdExtend
        strSel = .Text
        End With
        Debug.Print strSel
    With XReg
            .Pattern = "(\d*?)\.(\d*?)\."
        Set xMatches = .Execute(strSel)
    End With
        If Not xMatches.Count = 1 Then
            ActiveDocument.Bookmarks.Add Name:="NSN" & Replace(NSN, "-", "")
        xlWksht.Hyperlinks.Add Anchor:=xlWksht.Range("B" & cRow), _
                               Address:=ActiveDocument.FullName & "#NSN" & Replace(NSN, "-", ""), _
                               ScreenTip:="IPD for NSN " & NSN, _
                               TextToDisplay:="Link to IPD"
        Debug.Print "Added bookmark!"
    End If
    Next
    xlWkbk.Save
    xlWkbk.Close
Set xlApp = Nothing
Set xlWkbk = Nothing
Set xlWksht = Nothing
Set xlRange = Nothing
End Sub