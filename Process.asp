<%

' Retrieve values from form fields and set as variables
imgOne = Request.Form("Image One")
imgTwo = Request.Form("Image Two")
imgThree = Request.Form("Image Three")
imgFour = Request.Form("Image Four")
imgFive = Request.Form("Image Five")
imgSix = Request.Form("Image Six")
imgSeven = Request.Form("Image Seven")

' Create the AspEmail message object
Set Mail = Server.CreateObject("Persits.MailSender")

' Set the from Name and E-Mail address using values retrieved from the form
Mail.From = "xhobobobbbyx@gmail.com"
Mail.FromName = "Andrew Wang"

' Add the e-mail recipient address - replace values within the quotes with your own
Mail.AddAddress "congrenw@andrew.cmu.edu"

' Set the subject for the e-mail
Mail.Subject = "MTurk Info"

' Create a string called bodytxt and build it line by line using values from the form
Bodytxt = "Details of Form submission :" & VbCrLf & VbCrLf
Bodytxt = Bodytxt & "Image One : " & imgOne & VbCrLf
Bodytxt = Bodytxt & "Image Two : " & imgTwo & VbCrLf
Bodytxt = Bodytxt & "Image Three : " & imgThree & VbCrLf
Bodytxt = Bodytxt & "Image Four : " & imgFour & VbCrLf
Bodytxt = Bodytxt & "Image Five : " & imgFive & VbCrLf
Bodytxt = Bodytxt & "Image Six : " & imgSix & VbCrLf
Bodytxt = Bodytxt & "Query Seven : " & imgSeven

' Set body text for the e-mail to the Bodytxt string we built
Mail.Body = Bodytxt



' Mail is sent - tidy up and delete the AspEmail message object
Set Mail = Nothing

%>