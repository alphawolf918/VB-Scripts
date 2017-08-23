Sub Extract_NSN_List()
    'This procedure performs the following steps:
    '(1)Open spreadsheet containing list of NSN's in column B
    '(2)Loop through NSN's, and search for each one in Word Document
    '(3)When NSN is found, copy line containing item#, NSN, and U/I to spreadsheet column D

    Dim lrow As Integer
    Dim xlApp As Object, xlWkbk As Object, xlWksht As Object
Set xlApp = GetObject(, "Excel.Application")
'Specify Excel file to open and search here(full path with file name)
Set xlWkbk = xlApp.Workbooks.Open("[REDACTED]")
xlWkbk.Activate
'Change Sheet1 if sheet name is different
Set xlWksht = xlWkbk.Worksheets("Sheet1")
'Set last row of range
lrow = 1479
    Dim xlRange As Excel.Range
'Set range of cells with lookup values (must be in a single column)
Set xlRange = xlWksht.Range("C2:C" & lrow)
Dim NSN As String  'Lookup string
    Dim strNSNInfo As String   'Return string
    Dim cRow As Integer   'Current row of spreadsheet
    cRow = 0
    '(B) Loop through NSN's in worksheet, find in Word Document, add bookmark at beginning of line, add hyperlink to bookmark in Excel.
    For Each Cell In xlRange
        Selection.HomeKey Unit:=wdStory
NSN = Cell.Value
        Debug.Print "Copying NSN " & NSN & "from cell " & Cell.Address
cRow = Cell.Row
        Selection.find.ClearFormatting
        With Selection.find
            .Text = NSN
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        Selection.find.Execute
        '(OPTION B)-- Adding bookmarks to document and hyperlinks to spreadsheet
        '--Uncomment below here to activate this option
        Selection.HomeKey Unit:=wdLine
    ActiveDocument.Bookmarks.Add Name:="NSN" & Replace(NSN, "-", "")
    xlWksht.Hyperlinks.Add Anchor:=xlWksht.Range("B" & cRow), _
        Address:=ActiveDocument.FullName & "#NSN" & Replace(NSN, "-", ""), _
        ScreenTip:="IPD for NSN " & NSN, _
        TextToDisplay:="Link to IPD"
'--End section here
    Next
    xlWkbk.Save
    xlWkbk.Close
Set xlApp = Nothing
Set xlWkbk = Nothing
Set xlWksht = Nothing
Set xlRange = Nothing
End Sub