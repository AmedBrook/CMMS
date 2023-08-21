Sub WhatsApp()

    Dim Contact As String
    Dim Message As String 
    Dim Obj As New DataObject 'Creating and object of type DataObject



    Contact = Sheets("Requiste").Range("A6").Value 'Creating the contact variable of personnel to whome will be sending the notification
    Message = Sheets("Requiste").Range("B6").Value ' Creating the message variable of the notification text

    .SetText Message 
    Obj.PutInClipboard 

    ActiveWorkbook.FollowHyperlink "https://wa.me/" & Contact ' Basically sending the notif to what ever in the "" string and the contact variable

    
    'Setting the wait time to establish the connexion

    Application.Wait (Now + TimeValue("00:00:20")) 'setting 20 seconds before retry to send again
    Call SendKeys("^v", True)
    Application.Wait (Now + TimeValue("00:00:10")) 'setting 10 seconds before retry to send again
    Call SendKeys("~", True)
    Application.Wait (Now + TimeValue("00:00:05")) 'setting 5 seconds before retry to send again



End Sub