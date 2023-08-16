
'Accessing maintenance data archive by clicking on Data-archive button
Private Sub Data_ArchiveButton_Click()

Dim PSW As String

PSW = InputBox(" Saisir le Mot de passe!", "NIM Maintenance GMAO")
If PSW = "######" Then                    'with ###### is the choosen password
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



'switching to maintenance task request tab by clicking on Demand button
Private Sub Demmand_Button_Click()
task_userform.Hide
Demande_Userform.Show
End Sub

'Accessing the maintenance KPI tab by clicking on KPI-Analysis button
Private Sub KPI_CommandButton_Click()
task_userform.Hide
frmCharts.Show
End Sub

'Tracing a maintenance task by clicking on tracer une intervention
Private Sub TRACER_UNE_INTERVENTION_Button_Click()
task_userform.Hide
Tracer_UserForm.Show
End Sub

Private Sub UserForm_Click()
End Sub
