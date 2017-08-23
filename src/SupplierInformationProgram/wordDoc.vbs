'Initialization of the form. This is where the
'program will start.
Sub init()
    supplierForm.Show
End Sub

'Meat of the program. Passes in supplier ID
'as a string by value. Called from the supplierForm
'user form.
Sub GetSetDocVars(ByVal suppID As String)
    
    'Define all variables up top. 'Variant' is kind
    'of a superclass for VB, it seems.
    Dim oVars As Variant
    Set oVars = ActiveDocument.Variables
    
    'Variable for storing the query.
    Dim strQuery As String
    
    'Doc variables.
    Dim prpCompanyName As String
    Dim prpSupplierID As String
    Dim prpMinOrder As String
    Dim prpPrepaidFreight As String
    Dim prpSalesRep As String
    Dim prpRepAgency As String
    Dim prpOther As String
    Dim prpStockReq As String
    Dim prpPayDiscounts As String
    
    'Open an ActiveX Database Object Database connection.
    Dim con As ADODB.Connection
    
    'Rows get stored here.
    Dim rs As ADODB.Recordset
    
    'This is where the information for the data source is stored.
    Dim sConString As String
    
    'Set the connection and recordset variables.
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    'Connection string. All that's needed is a data source named P21.
    sConString = "Data Source=Prophet21;Server=giga-p21sql;Database=Prophet21;Trusted_Connection=True;"
    
    'This "resets" all variable values. If this is omitted, the program
    'will complain and say that the variables already exist.
    oVars("prpCompanyName").Value = ""
    oVars("prpSupplierID").Value = ""
    oVars("prpMinOrder").Value = ""
    oVars("prpPrepaidFreight").Value = ""
    oVars("prpDate").Value = ""
    oVars("prpSalesRep").Value = ""
    oVars("prpRepAgency").Value = ""
    oVars("prpOther").Value = ""
    oVars("prpStockReq").Value = ""
    oVars("prpPayDiscounts").Value = ""
    
    'Set to blank values.
    oVars("prpCompanyName").Value = " "
    oVars("prpSupplierID").Value = " "
    oVars("prpMinOrder").Value = " "
    oVars("prpPrepaidFreight").Value = " "
    oVars("prpDate").Value = " "
    oVars("prpSalesRep").Value = " "
    oVars("prpRepAgency").Value = " "
    oVars("prpOther").Value = " "
    oVars("prpStockReq").Value = " "
    oVars("prpPayDiscounts").Value = " "
    
    If suppID = vbNullString Then
        suppID = 0
    End If
    
    'Open connection.
    con.Open sConString
    
    'SQL query using passed supplier ID. Additionally, if the target_value
    'is NULL, it displays 0.00 in its place (COALESCE).
    strQuery = "SELECT " & _
               "p21_supplier_view.supplier_id AS 'Supplier ID'," & _
               "p21_supplier_view.supplier_name AS 'Supplier Name'," & _
               "p21_supplier_view.freight_target_value AS 'Prepaid Freight'," & _
               "COALESCE(p21_supplier_view.target_value, 0.00) AS 'Min Order' " & _
               "FROM p21_supplier_view " & _
               "WHERE p21_supplier_view.supplier_id = '" & suppID & "';"
    
    'Run the query and store the results in the record set.
    Set rs = con.Execute(strQuery)
    
    'If not the (E)nd (O)f the (F)ile (the record set), keep looping.
    'Sets VB variables equal to database values if a match is found,
    'or displays a message box if no records are found.
    If Not rs.EOF Then
       prpSupplierID = rs.Fields("Supplier ID")
       prpCompanyName = rs.Fields("Supplier Name")
       prpMinOrder = rs.Fields("Min Order")
       prpPrepaidFreight = rs.Fields("Prepaid Freight")
       rs.Close
    Else
        MsgBox "No records found.", vbOKOnly
    End If
    
    If prpMinOrder = vbNullString Then
        prpMinOrder = " "
    End If
    
    If prpPrepaidFreight = vbNullString Then
        prpPrepaidFreight = " "
    End If
    
    'If the connection is still open, close it.
    If CBool(con.State And adStateOpen) Then con.Close
    
    oVars("prpCompanyName").Value = ""
    oVars("prpSupplierID").Value = ""
    oVars("prpMinOrder").Value = ""
    oVars("prpPrepaidFreight").Value = ""
    
    'Set doc variables equal to VB variables (which in turn contain
    'the database values).
    oVars.Add Name:="prpCompanyName", Value:=prpCompanyName
    oVars.Add Name:="prpSupplierID", Value:=prpSupplierID
    oVars.Add Name:="prpMinOrder", Value:=prpMinOrder
    oVars.Add Name:="prpPrepaidFreight", Value:=prpPrepaidFreight
    
    'Sets the current date.
    oVars("prpDate").Value = ""
    oVars.Add Name:="prpDate", Value:=Date
    
    'New or Existing Supplier
    If supplierForm.chkNew.Value = True Then
        FormFields.Item("Check7").CheckBox.Value = True
    ElseIf supplierForm.chkExisting.Value = True Then
        FormFields.Item("Check8").CheckBox.Value = True
    End If
    
    'Affiliations
    If supplierForm.chkNetPlus.Value = True Then
        FormFields.Item("netPlus").CheckBox.Value = True
    End If
    If supplierForm.chkOrgill.Value = True Then
        FormFields.Item("orgill").CheckBox.Value = True
    End If
    
    'Sales Rep Info
    If supplierForm.chkRegional.Value = True Then
        FormFields.Item("Check2").CheckBox.Value = True
    End If
    If supplierForm.chkRepAgency.Value = True Then
        FormFields.Item("Check1").CheckBox.Value = True
        If supplierForm.txtRep.Value = vbNullString Then
            prpRepAgency = " "
        Else
            prpRepAgency = supplierForm.txtRep.Value
        End If
    Else
        prpAgency = " "
    End If
    
    'Rep Agency
    oVars("prpRepAgency").Value = ""
    If prpRepAgency = "" Or prpRepAgency = vbNullString Then
        prpRepAgency = " "
    End If
    oVars.Add Name:="prpRepAgency", Value:=prpRepAgency
    
    'Grab the entered sales rep text from the form.
    If prpSalesRep = "" Or prpSalesRep = vbNullString Then
        prpSalesRep = " "
    Else
        prpSalesRep = supplierForm.txtSalesRepName.Value
    End If
    oVars("prpSalesRep").Value = ""
    oVars.Add Name:="prpSalesRep", Value:=prpSalesRep
    
    'Supplier Type
    If supplierForm.chkManu.Value = True Then
        FormFields.Item("Check3").CheckBox.Value = True
    ElseIf supplierForm.chkDistr.Value = True Then
        FormFields.Item("Check4").CheckBox.Value = True
    ElseIf supplierForm.chkOth.Value = True Then
        FormFields.Item("Check6").CheckBox.Value = True
        If supplierForm.txtOth.Value = vbNullString Then
            prpOther = " "
        Else
            prpOther = supplierForm.txtOth.Value
        End If
    Else
        prpOther = " "
    End If
    
    'Other
    If prpOther = "" Or prpOther = vbNullString Then
        prpOther = " "
    End If
    oVars("prpOther").Value = ""
    oVars.Add Name:="prpOther", Value:=prpOther
    
    'Stocking Requirement
    If supplierForm.txtStockReq.Value = vbNullString Then
        prpStockReq = " "
    Else
        prpStockReq = supplierForm.txtStockReq.Value
    End If
    If prpStockReq = "" Or prpStockReq = vbNullString Then
        prpStockReq = " "
    End If
    oVars("prpStockReq").Value = ""
    oVars.Add Name:="prpStockReq", Value:=prpStockReq
    
    'Payment Discounts
    If supplierForm.txtDiscounts.Value = vbNullString Then
        prpPayDiscounts = " "
    Else
        prpPayDiscounts = supplierForm.txtDiscounts.Value
    End If
    oVars("prpPayDiscounts").Value = ""
    oVars.Add Name:="prpPayDiscounts", Value:=prpPayDiscounts
    
    'Open the Excel link and pass in the supplier ID.
    updateExcelLinks (suppID)
    
    'Clean up.
    Set con = Nothing
    Set rs = Nothing
    
    'Run the update method.
    updateAllFields
End Sub

Private Sub updateExcelLinks(ByVal suppID As String)
    Dim linkedChart As workBook
    Set linkedChart = Workbooks.Open("S:\Matt's Team Stuff\SRM\Supplier Info Sheets" _
                      & "\P21-Supplier Purchase History.xlsx", 1)
    Sheets("PT-PO Totals").Select
    Worksheets("PT-PO Totals").Range("B1").Value = suppID
    Sheets("Purchase Totals").Select
    
    linkedChart.Close (1)
    
    Set linkedChart = Nothing
End Sub

'Forces the document to update all DocVars to show their
'new values.
Private Sub updateAllFields()
Dim oStyRng As Word.Range
Dim iLink As Long
  iLink = ActiveDocument.Sections(1).Headers(1).Range.StoryType
  For Each oStyRng In ActiveDocument.StoryRanges
    Do
      oStyRng.Fields.Update
      Set oStyRng = oStyRng.NextStoryRange
    Loop Until oStyRng Is Nothing
  Next
End Sub

Private Sub Document_Open()
    Dim oVars As Variant
    Set oVars = wordDoc.Variables
    oVars("prpCompanyName").Value = ""
    oVars("prpSupplierID").Value = ""
    oVars("prpMinOrder").Value = ""
    oVars("prpPrepaidFreight").Value = ""
    oVars("prpDate").Value = ""
    oVars("prpSalesRep").Value = ""
    oVars("prpRepAgency").Value = ""
    oVars("prpOther").Value = ""
    oVars("prpStockReq").Value = ""
    oVars("prpPayDiscounts").Value = ""
    
    oVars("prpCompanyName").Value = " "
    oVars("prpSupplierID").Value = " "
    oVars("prpMinOrder").Value = " "
    oVars("prpPrepaidFreight").Value = " "
    oVars("prpDate").Value = " "
    oVars("prpSalesRep").Value = " "
    oVars("prpRepAgency").Value = " "
    oVars("prpOther").Value = " "
    oVars("prpStockReq").Value = " "
    oVars("prpPayDiscounts").Value = " "
    oVars("prpDate").Value = ""
    oVars.Add Name:="prpDate", Value:=Date
    
    updateAllFields
    init
End Sub
