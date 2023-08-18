
Private Sub Annuler_Button_Click()
    'Initializing fields
    Demande_Userform.Interveneur_TextBox.Value = ""
    Demande_Userform.Cause_TextBox.Value = ""
    Demande_Userform.Corrective_OptionButton = False
    Demande_Userform.Preventive_OptionButton = False
    Demande_Userform.DateIntervention_TextBox.Value = ""
    Demande_Userform.HeureInter_TextBox.Value = ""
    Demande_Userform.temps_arrêt_éstimé_TextBox.Value = ""
    Demande_Userform.Machine_Zone_TextBox.Value = ""
    Demande_Userform.Piécesderechange_TextBox.Value = ""
End Sub


Private Sub Cause_TextBox_Change()
End Sub

Private Sub Demander_Button_Click()
    Dim Lastrow As Long
    Lastrow = WorksheetFunction.CountA(Sheets("Demandes").Range("A:A"))

    'Checking if all boxes are filled 
    If Demande_Userform.Interveneur_TextBox.Value = "" Or Demande_Userform.Cause_TextBox.Value = "" Or (Demande_Userform.Corrective_OptionButton = False And Demande_Userform.Preventive_OptionButton = False) Or Demande_Userform.DateIntervention_TextBox.Value = "" Or Demande_Userform.HeureInter_TextBox.Value = "" Or Demande_Userform.temps_arrêt_éstimé_TextBox.Value = "" Or Demande_Userform.Machine_Zone_TextBox.Value = "" Or Demande_Userform.Piécesderechange_TextBox.Value = "" Then
        MsgBox " Il faut Remplir toutes les cases ! "
        Else
            Sheets("Demandes").Cells(Lastrow + 1, 1).Value = Lastrow
            Sheets("Demandes").Cells(Lastrow + 1, 2).Value = Demande_Userform.Interveneur_TextBox.Value

        If Demande_Userform.Corrective_OptionButton = True Then
        Sheets("Demandes").Cells(Lastrow + 1, 3).Value = "Corrective"
        Else
        Sheets("Demandes").Cells(Lastrow + 1, 3).Value = "Préventive"
        End If

        'Incrementing rows 
        Sheets("Demandes").Cells(Lastrow + 1, 4).Value = Demande_Userform.Cause_TextBox.Value
        Sheets("Demandes").Cells(Lastrow + 1, 5).Value = Demande_Userform.DateIntervention_TextBox.Value
        Sheets("Demandes").Cells(Lastrow + 1, 6).Value = Demande_Userform.HeureInter_TextBox.Value
        Sheets("Demandes").Cells(Lastrow + 1, 7).Value = Demande_Userform.temps_arrêt_éstimé_TextBox.Value
        Sheets("Demandes").Cells(Lastrow + 1, 8).Value = Demande_Userform.Machine_Zone_TextBox.Value
        Sheets("Demandes").Cells(Lastrow + 1, 9).Value = Demande_Userform.Piécesderechange_TextBox.Value

        'Calling the subrouting for remote notifying, here I am using WhatsApp as an example
        Call WhatsApp
        MsgBox "demand has been transfered succefully ! "
        Call Annuler_Button_Clic

    End If

End Sub



Private Sub Interveneur_TextBox_Change()
End Sub

Private Sub Menu_Button_Click()
    Demande_Userform.Hide
    task_userform.Show
End Sub


Private Sub Piécesderechange_TextBox_Change()
End Sub


Private Sub temps_arrêt_éstimé_TextBox_Change()
End Sub

Private Sub UserForm_Click()
End Sub
