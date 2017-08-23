Private Sub TestPDFedit()
    Dim OpDoc As Object
    Dim annot As Object
    Dim page As Acrobat.AcroPDPage
    Dim intpoint(1) As Integer
    Dim Today As String
    Dim strPath As String
    strPath = "[REDACTED]"
    Dim strExtension As String
    strExtension = ".pdf"
    Dim strFullPath As String
    strFullPath = strPath + strExtension
    Dim AcroApp As Acrobat.CAcroApp
    Dim MIL129DOC As Acrobat.CAcroPDDoc
    Dim props As Object
 Set AcroApp = CreateObject("AcroExch.App")
 Set MIL129DOC = CreateObject("AcroExch.PDDoc")
 Today = Format(Now(), "MMDDYY")
    Dim intpopupRect(3) As Integer
    Dim jso As Object
    MIL129DOC.Open(strPath)
 Set jso = MIL129DOC.GetJSObject
 Dim field As Object
    Dim i As Long
    Dim strNewPath As String
    strNewPath = Replace(strPath, ".pdf", "") & "-" & Today & strExtension


    If Not jso Is Nothing Then

        ' From API Reference:
        '
        ' * Origin Point = Lower left corner
        '
        ' * The positive x-axis points to
        '   the right from the origin.
        '
        ' * The positive y-axis moves up
        '   from the origin.
        '
        ' * The length of one unit along
        '   both X and Y axes is 1/72 inch.

        intpopupRect(0) = 540 ' x upper right
        intpopupRect(1) = 130 ' y upper right
        intpopupRect(2) = 385 ' x lower left (Origin Point)
        intpopupRect(3) = 265 ' y lower left

Set field = jso.addField("textField", "text", 3, intpopupRect)
field.textSize = 10
        field.Value = CStr(DateValue(Now()))
        jso.flattenPages   'Pushes objects added on top of document into document

    End If
    i = MIL129DOC.Save(PDSaveCopy, strNewPath)
    'AcroApp.MenuItemExecute ("SaveAs")
    MIL129DOC.Close

Set pdDoc = Nothing

MsgBox "Done"

End Sub