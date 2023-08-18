
Private Sub Ajouté_CommandButton_Click()

    Dim Lastrow As Long
    Lastrow = WorksheetFunction.CountA(Sheets("Intervention").Range("A:A"))

    If Tracer_UserForm.RéfIntervention_TextBox.Value = "" Or Tracer_UserForm.Machine_TextBox.Value = "" Or (Tracer_UserForm.Corrective_OptionButton = False And Tracer_UserForm.Preventive_OptionButton = False And Tracer_UserForm.Conditionnlle_OptionButton = False) Or Tracer_UserForm.DateIntervention_TextBox.Value = "" Or Tracer_UserForm.Durée_TextBox.Value = "" Or Tracer_UserForm.Outilsutilisés_TextBox.Value = "" Or Tracer_UserForm.Périodicité_ComboBox.Value = "" Or MatiéresConsommés_TextBox.Value = "" Then
        MsgBox " Il faut Remplir toutes les cases ! "

    Else
        Sheets("Intervention").Cells(Lastrow + 1, 1).Value = Lastrow
        Sheets("Intervention").Cells(Lastrow + 1, 2).Value = Tracer_UserForm.Machine_TextBox.Value

    If Tracer_UserForm.Corrective_OptionButton = True Then
        Sheets("Intervention").Cells(Lastrow + 1, 3).Value = "Corrective"
 
    ElseIf Tracer_UserForm.Conditionnlle_OptionButton = True Then
        Sheets("Intervention").Cells(Lastrow + 1, 3).Value = "Conditionnelle"

    Else
        Sheets("Intervention").Cells(Lastrow + 1, 3).Value = "Preventive"
        MsgBox "Data has been updated succefully"

    End If
    End If

    Sheets("Intervention").Cells(Lastrow + 1, 4).Value = Tracer_UserForm.DateIntervention_TextBox.Value
    Sheets("Intervention").Cells(Lastrow + 1, 5).Value = Tracer_UserForm.Durée_TextBox.Value
    Sheets("Intervention").Cells(Lastrow + 1, 6).Value = Tracer_UserForm.Outilsutilisés_TextBox.Value
    Sheets("Intervention").Cells(Lastrow + 1, 7).Value = Tracer_UserForm.Périodicité_ComboBox.Value
    Sheets("Intervention").Cells(Lastrow + 1, 8).Value = Tracer_UserForm.DescriptionNotes_TextBox.Value
    Sheets("Intervention").Cells(Lastrow + 1, 9).Value = Tracer_UserForm.Piécesderechange_TextBox.Value
    Sheets("Intervention").Cells(Lastrow + 1, 10).Value = Tracer_UserForm.MatiéresConsommés_TextBox.Value
    Call Annulé_CommandButton_Click

End Sub


Private Sub Annulé_CommandButton_Click()

    Tracer_UserForm.RéfIntervention_TextBox.Value = ""
    Tracer_UserForm.Machine_TextBox.Value = ""
    Tracer_UserForm.Corrective_OptionButton = False
    Tracer_UserForm.Preventive_OptionButton = False
    Tracer_UserForm.DateIntervention_TextBox.Value = ""
    Tracer_UserForm.Durée_TextBox.Value = ""
    Tracer_UserForm.Outilsutilisés_TextBox.Value = ""
    Tracer_UserForm.Périodicité_ComboBox.Value = ""
    DescriptionNotes_TextBox.Value = ""
    Piécesderechange_TextBox.Value = ""
    MatiéresConsommés_TextBox.Value = ""
End Sub



Private Sub Data_CommandButton_Click()

    Dim PSW As String

    PSW = InputBox(" Saisir le Mot de passe!", "NIM Maintenance GMAO")
    If PSW = "####" Then
        Application.Visible = True
        Sheets("Intervention").Visible = True
        Sheets("Intervention").Activate
        Tracer_UserForm.Hide

    ElseIf PSW = "" Then
        MsgBox " Mot de passe est incorrect ! ", vbOKCancel, " NIM Maintenance GMAO"

    Else
        MsgBox " Mot de passe est incorrect ! ", vbOKCancel, " NIM Maintenance GMAO"
    End If

End Sub


Private Sub DescriptionNotes_TextBox_Change()
End Sub

Private Sub Label12_Click()
End Sub

Private Sub Menu_Button_Click()
    Tracer_UserForm.Hide
    task_userform.Show
End Sub


Private Sub Périodicité_ComboBox_Change()
End Sub

Private Sub UserForm_Terminate()
    ActiveWorkbook.Save
    Application.Quit
End Sub
