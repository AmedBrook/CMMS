Sub WhatsApp()

    Dim Contact As String
    Dim Message As String
    Dim Obj As New DataObject



    Contact = Sheets("Requiste").Range("A6").Value
    Message = Sheets("Requiste").Range("B6").Value

    .SetText Message
    Obj.PutInClipboard

    ActiveWorkbook.FollowHyperlink "https://wa.me/" & Contact

    Application.Wait (Now + TimeValue("00:00:20"))
    Call SendKeys("^v", True)

    Application.Wait (Now + TimeValue("00:00:10"))
    Call SendKeys("~", True)
    Application.Wait (Now + TimeValue("00:00:05"))



End Sub