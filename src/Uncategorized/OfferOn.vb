Sub Extract_BOE()

Dim docBOE, docCurr As Object
Dim intFound As Integer
Dim startIPD As Integer
Dim ipdIndex As Integer
Dim endIPD As Integer
Dim rngIPDstart, rngIPDend As Range
Dim iCount As Integer
Dim numBookMarks As Integer
Set docCurr = ActiveDocument
Set docBOE = Documents.Add(, , , 1)
docCurr.Activate

Selection.HomeKey Unit:=wdStory
Selection.Find.ClearFormatting
    With Selection.Find
        .ClearFormatting
        .Text = "Offering on"
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
    End With
    
    Do While Selection.Find.Execute(Forward:=True) And iCount < 1000
        iCount = iCount + 1
        numBookMarks = ActiveDocument.Bookmarks.Count
        If Selection.Find.Found Then
            startIPD = Selection.Range.PreviousBookmarkID
            endIPD = startIPD + 1
            Debug.Print
            Debug.Print ActiveDocument.Bookmarks(startIPD).Name
            Debug.Print ActiveDocument.Bookmarks(endIPD).Name
            
            Set rngIPDstart = ActiveDocument.Bookmarks(startIPD).Range
            Set rngIPDend = ActiveDocument.Bookmarks(endIPD).Range
            ActiveDocument.Range(rngIPDstart.Start, rngIPDend.End).Select
            With Selection
                .Select
                .Copy
            End With
            Documents(docBOE).Activate
            Selection.PasteAndFormat (wdFormatOriginalFormatting)
            Documents(docCurr).Activate
            With Selection.Find
                .ClearFormatting
                .Text = "Offering on"
                .Wrap = wdFindContinue
                .Format = False
                .Forward = True
                .MatchCase = False
                .Execute
            End With
         Else
            MessageBox.Show ("The text was not found.")
            Exit Do
        End If
    Loop
End Sub