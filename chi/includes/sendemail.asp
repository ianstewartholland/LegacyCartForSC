<!--#include file ="WriteLogEntry.asp"-->
<%
'
' send email function to send automated emails from localhost using CDOSYS
' try to log failed messages if there are any - remember to check the log occasionally
'
Public Function SendMail(strTo, _
			strFrom, _
			strSubject, _
			strBody)

	' Send mail using the sentex mail server.
	On Error Resume Next
	Set objNewMail = CreateObject("CDO.Message")
	objNewMail.Subject = strSubject
	objNewMail.From = strFrom
	objNewMail.To = strTo
	objNewMail.TextBody = strBody
	
	'Name or IP of remote SMTP server
	objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	objNewMail.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
	objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objNewMail.Configuration.Fields.Update

	objNewMail.Send
	Set objNewMail = Nothing
	SendMail = True

	If Err.Number <> 0 Then
	    WriteLogEntry("**************************Please do not ignore this message**************************")
	    WriteLogEntry("Important! auto-email failed on " & Time & ", " & Date)
	    WriteLogEntry("Error code: " & Err.Number)
	    WriteLogEntry("Email was supposed to be sent to: " & strTo)
	    WriteLogEntry("Original email body was: ")
	    WriteLogEntry(strBody)
	    WriteLogEntry("*************************************************************************************")
	    SendMail = False
	End If

End Function
%>