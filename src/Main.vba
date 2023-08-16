

Private Sub Data_ArchiveButton_Click()

Dim PSW As String

PSW = InputBox(" Saisir le Mot de passe!", "NIM Maintenance GMAO")
If PSW = "1234" Then
Application.Visible = True
Sheets("Intervention").Visible = True
Sheets("Intervention").Activate
task_userform.Hide

ElseIf PSW = "" Then

MsgBox " Mot de passe est incorrect ! ", vbOKCancel, " NIM Maintenance GMAO"

Else

MsgBox " Mot de passe est incorrect ! ", vbOKCancel, " NIM Maintenance GMAO"


End If
End Sub

Private Sub Demmand_Button_Click()

task_userform.Hide
Demande_Userform.Show

End Sub


Private Sub KPI_CommandButton_Click()


task_userform.Hide
frmCharts.Show

End Sub

Private Sub TRACER_UNE_INTERVENTION_Button_Click()
task_userform.Hide
Tracer_UserForm.Show

End Sub

Private Sub UserForm_Click()

End Sub
