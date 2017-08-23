Sub CreateBookmarks()

Dim strSel As String
Dim XReg As RegExp
Dim xMatches As MatchCollection
Dim intAmount As Long
Dim intRangeStart As Long
Dim intRangeEnd As Long

intAmount = 0

Dim p As Paragraph

For Each p In ActiveDocument.Paragraphs
    Debug.Print p.Range.Text
    strSel = p.Range.Text
    Set XReg = New RegExp

    With XReg
        .IgnoreCase = False
        .Global = False
        .Pattern = "^T(\d*?)\.(\d*?)"
        Set xMatches = .Execute(strSel)
    End With
    If xMatches.Count >= 1 Then
        Debug.Print "Match found!"
        intRangeStart = p.Range.Start
        intRangeEnd = intRangeStart
        ActiveDocument.Range(intRangeStart, intRangeEnd).Select
        ActiveDocument.Bookmarks.Add Name:="T" & intAmount
        Debug.Print "Added bookmark!"
        intAmount = intAmount + 1
    End If
Next p

End Sub