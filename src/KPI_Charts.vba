'Closing screen subroutine
Private Sub Fermer_CommandButton_Click()
    Unload Me
End Sub

'Screen switch
Private Sub Menu_CommandButton_Click()
    frmCharts.Hide
    task_userform.Show
End Sub

'Visulaizing down time of caused by equipement maintenance.
Private Sub Opt_Temps_arrêt_Click()
    Call ChangeChart("Maintenance down time")
End Sub

'Visualizing the share of each type of maintenance.
Private Sub Opt_Typemaintenance_Click()
    Call ChangeChart("Maintenance type")
End Sub

Private Sub UserForm_Click()
End Sub
