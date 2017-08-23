Private Sub btnCancel_Click()
   Unload supplierForm
End Sub

Private Sub btnOk_Click()
    supplierForm.Hide
    Dim strSuppID As Variant
    If supplierForm.txtSuppID.Value = vbNullString Then
        MsgBox "No supplier ID input given!", vbCritical
        Unload supplierForm
        wordDoc.init
    Else
        strSuppID = supplierForm.txtSuppID.Value
    End If
    wordDoc.GetSetDocVars strSuppID
    Unload supplierForm
End Sub

Private Sub chkOth_Change()
    If txtOth.Enabled Then
        txtOth.BackColor = &H80000016
        txtOth.Value = ""
        txtOth.Enabled = False
    End If
End Sub

Private Sub chkOth_Click()
    txtOth.BackColor = &H80000014
    txtOth.Enabled = True
End Sub

Private Sub chkRepAgency_Change()
    If txtRep.Enabled Then
        txtRep.BackColor = &H80000016
        txtRep.Value = ""
        txtRep.Enabled = False
    End If
End Sub

Private Sub chkRepAgency_Click()
    If chkRepAgency.Value = True Then
        txtRep.BackColor = &H80000005
        txtRep.Enabled = True
    Else
        txtRep.BackColor = &H80000016
        txtRep.Value = ""
        txtRep.Enabled = False
    End If
End Sub

Private Sub UserForm_Initialize()
    txtSuppID.SetFocus
End Sub
