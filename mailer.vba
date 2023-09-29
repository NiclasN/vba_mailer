Sub SendEmailWithFooter()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strBody As String
    Dim strFooter As String
    Dim i As Long
    Dim lastRow As Long
    
    ' Get the last row of data in the sheet
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Create Outlook application and mail item
    Set OutApp = CreateObject("Outlook.Application")
    
    For i = 2 To lastRow ' Start from row 2 assuming row 1 is header
        ' Create the email body with HTML
        strBody = "<p>Dear " & Cells(i, 4).Value & "</p>" & _
                  "<p>" & Cells(i, 5).Value & "</p>"
        
        strFooter = "<footer style='border-top: 1px solid #ddd; padding-top: 10px;'>" & _
                      "<p>Best Regards,</p>" & _
                      "<p>" & Cells(i,6).Value & "</p>" & _
                      "<p>The Information Lab Deutschland GmbH</p>" & _
                      "<img src='https://www.theinformationlab.de/wp-content/uploads/2021/11/TIL-Logo_neu_390x157-pixel.png' href=https://www.theinformationlab.de alt='The Information Lab Deutschland GmbH'>" & _
                  "</footer>"
        
        Set OutMail = OutApp.CreateItem(0)
        
        With OutMail
            .To = Cells(i, 1).Value
            .CC = Cells(i, 2).Value
            .BCC = Cells(i, 3).Value
            .Subject = "Test Email with Footer"
            .HTMLBody = strBody & strFooter
            .Display ' Use .Send to send the email directly
        End With
        
        ' Clean up
        Set OutMail = Nothing
    Next i
    
    Set OutApp = Nothing

End Sub