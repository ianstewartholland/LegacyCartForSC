<%
	'
	'check password for regular login
	'
	If Session("PasswordOK") <> "Yes" Then Response.Redirect "/chi/protected/chibackend.aspx"
	
	'
	'check password level
	'
	isLocked = True
	If Session("PasswordLevel") = "Owner" Or Session("PasswordLevel") = "Admin" Or Session("PasswordLevel") = "Bookkeeper" Then isLocked = False
%>