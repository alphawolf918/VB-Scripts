Sub CreateBookmarks()

    Dim lrow As Integer
    Dim xlApp As Object, xlWkbk As Object, xlWksht As Object
Set xlApp = GetObject(, "Excel.Application")
Dim strSel As String
    Dim XReg As RegExp
    Dim xMatches As MatchCollection
    Dim xMatch As Match
    Dim p As Paragraph
    Dim NSN As String
Set xlWkbk = xlApp.Workbooks.Open("[REDACTED]")
xlWkbk.Activate

Set xlWksht = xlWkbk.Worksheets("Sheet1")

lrow = 218

    Dim xlRange As Excel.Range

Set xlRange = xlWksht.Range("A1:A" & lrow)

Dim cRow As Integer
    cRow = 0

    For Each Cell In xlRange
        ' Selection.HomeKey Unit:=wdStory
        NSN = Cell.Value
        cRow = Cell.Row
        With Selection.Find
            .Text = NSN
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute
        End With
        If Selection.Find.Found Then
            With Selection
                strSel = .Text
            End With
            Debug.Print strSel
        ActiveDocument.Bookmarks.Add Name:="NSN" & Replace(NSN, ".", "")
        xlWksht.Hyperlinks.Add Anchor:=xlWksht.Range("B" & cRow), _
                               Address:=ActiveDocument.FullName & "#NSN" & Replace(NSN, ".", ""), _
                               ScreenTip:="IPD for NSN " & NSN, _
                               TextToDisplay:="Link to IPD"
        Debug.Print "Added bookmark!"
    End If
    Next

Set xlApp = Nothing
Set xlWkbk = Nothing
Set xlWksht = Nothing
Set xlRange = Nothing

End Sub