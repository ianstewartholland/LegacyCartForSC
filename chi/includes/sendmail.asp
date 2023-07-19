

<%

Public Function SendMail(strTo, _
						strFrom, _
						strSubject, _
						strBody, _
						strUsername, _
						strPassword)
						
' Send mail using the sentex mail server.
	On Error Resume Next
	Set objNewMail = CreateObject("CDO.Message")
	objNewMail.Subject = strSubject
	objNewMail.From = strFrom
	objNewMail.To = strTo
	objNewMail.TextBody = strBody

	objNewMail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
	'Name or IP of remote SMTP server
	'objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="vmail1.sentex.ca"
	'objNewMail.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	'objNewMail.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = strUsername
	'objNewMail.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strPassword
objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="localhost"
objNewMail.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25

	'Server port
	'objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=587
	objNewMail.Configuration.Fields.Update

	objNewMail.Send
	Set objNewMail = Nothing

	SendMail = True
	if Err.Number <> 0 then
		SendMail = False
	End If


End Function


%>