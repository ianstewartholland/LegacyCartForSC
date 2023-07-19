<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<%
'
' generic copy page to work with license, enterprise license, francelicense, francecompany and order pages
'
' <a target="_blank" href="/chi/includes/copy.asp?Type=EnterpriseLicense&ID=">Copy</a>&nbsp;&nbsp;&nbsp;
' <a href="/chi/includes/paste.asp?Type=EnterpriseLicense&ID=">Paste</a>&nbsp;&nbsp;&nbsp;
'	
	' retrieve values
	If Len(Request.QueryString("Type")) > 1 And Len(Request.QueryString("ID")) > 1 Then
		argType = Request.QueryString("Type")
		ID = Request.QueryString("ID")
        Else 
		Response.Redirect "/chi/cart/home.asp"
	End If

	' now check password for both regular and france
	If Left(LCase(argType), 6) = "france" Then
		If Session("PasswordFrance") <> "Yes" Then Response.Redirect "/chi/france/chipassword.asp"
	Else
%>
	    <!--#include virtual ="/chi/includes/pscheck.inc"-->
<%
	End If
%>

<html>

<head>
<title>Copy Address for <% Response.Write argType & ": " & ID %> - Computational Hydraulics Int.</title>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache"> <% 'force page refresh %>
<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<META HTTP-EQUIV="EXPIRES" CONTENT="0">
<!meta description="Consulting engineering firm specializing in urban stormwater quantity and quality management modeling esp SWMM">
<!meta keywords="SWMM USEPA SWMM4.4 PCSWMM stormwater management model modeling modelling BMP's catchbasin civil engineering channel computer conduit discharge hydraulics hydrology management modeling modelling NPDES non-pointsource pipe pollutants precipitation programming quality rain rainfall runoff sediment simulation storm stormwater streamflow transport urban water quality water resources engineering">
</head>

<body background="../cart/images/background.gif"><%
' 
' retrieve values from the database
'
Select Case argType
    Case "License", "FranceLicense", "EnterpriseLicense", "FranceCompany"
	'connect to database
	Set objRec = Server.CreateObject ("ADODB.Recordset")	
	strConnect = licensePath
	
	If argType = "EnterpriseLicense" Then 
	    objRec.Open "Enterprise", strConnect, 3, 3, 2
	    objRec.Filter = "EnterpriseLicense = '" & ID & "'"
	ElseIf argType = "FranceCompany" Then 
	    objRec.Open "France", strConnect, 3, 3, 2
	    objRec.Filter = "FranceLicense = '" & ID & "'"	
	Else
	    objRec.Open "License", strConnect, 3, 3, 2
	    objRec.Filter = "License = '" & ID & "'"
	End If
	
	objRec.MoveFirst
    
	Prefix = ObjRec("Prefix")
	FirstName = ObjRec("FirstName")
	LastName = ObjRec("LastName")
	Company = ObjRec("Company")
	StreetAddress = ObjRec("StreetAddress")
	StreetAddress2 = ObjRec("StreetAddress2")
	City = ObjRec("City")
	Province = ObjRec("Province")
	PostalCode = ObjRec("PostalCode")
	Email = ObjRec("Email")
	Country = ObjRec("Country")
	Telephone = ObjRec("Telephone")
	Fax = ObjRec("Fax")
	WebSite = ObjRec("WebSite")
	LicType = ObjRec("Type")
	
	'close database connection
	objRec.Close
	Set objRec = Nothing	
	
    Case "Transaction"
	'connect to database and load the information for this page
	Set objRec = Server.CreateObject ("ADODB.Recordset")	
	strConnect = transcationPath
	objRec.Open "[Transaction]", strConnect, 3, 3, 2
	objRec.Filter = "TransactionNumber = '" & ID & "'"
	objRec.MoveFirst
	
	Prefix = ObjRec("ClientPrefix")
	FirstName = ObjRec("ClientFirstName")
	LastName = ObjRec("ClientLastName")
	Company = ObjRec("ClientCompany")
	StreetAddress = ObjRec("ClientStreetAddress")
	StreetAddress2 = ObjRec("ClientStreetAddress2")
	City = ObjRec("ClientCity")
	Province = ObjRec("ClientProvince")
	Country = ObjRec("ClientCountry")
	PostalCode = ObjRec("ClientPostalCode")
	Email = ObjRec("ClientEmail")
	Telephone = ObjRec("ClientTelephone")
	Fax = ObjRec("ClientFax")
	OrderDate = ObjRec("OrderDate")	

	'close database connection	
	objRec.Close
	Set objRec = Nothing
	
    Case Else
	Response.Write "<font color = 'red'><b>Hey we do not have the information you requested.</b></font>"
	
End Select

	' write the value to the browser
    	Response.Write "Prefix=" & Prefix & "<br>"
    	Response.Write "FirstName=" & FirstName & "<br>"
    	Response.Write "LastName=" & LastName & "<br>"
    	Response.Write "Firm=" & Company & "<br>"
    	Response.Write "StreetAddress=" & StreetAddress & "<br>"
    	Response.Write "StreetAddress2=" & StreetAddress2 & "<br>"
    	Response.Write "City=" & City & "<br>"
    	Response.Write "Province=" & Province & "<br>"
    	Response.Write "ZipCode=" & PostalCode & "<br>"
    	Response.Write "Country=" & Country & "<br>"
    	Response.Write "Email=" & Email & "<br>"
    	Response.Write "Telephone=" & Telephone & "<br>"
    	Response.Write "Fax=" & Fax & "<br>"
	If argType <> "Transaction" Then 
	    Response.Write "WebSite=" & WebSite & "<br>"
	    Response.Write "Type=" & LicType & "<br>"
	Else
	    OrderDate = DatePart("yyyy", OrderDate) * 100 + DatePart("m", OrderDate)
	    Response.Write "LastContactDate=" & OrderDate & "<br>"
	    Response.Write "LastAddressUpdate=" & OrderDate & "<br>"
	End If
	
%></body>
&nbsp;</html>




 
 