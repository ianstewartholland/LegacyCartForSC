<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<%
'
' generic paste page to work with license, enterprise license, francelicense, francecompany and order pages
'
	' retrieve values
	If Len(Request.QueryString("Type")) > 1 And Len(Request.QueryString("ID")) > 1 Then
		argType = Request.QueryString("Type")
		ID = Request.QueryString("ID")
        Else 
		Response.Redirect "/chi/cart/home.asp"
	End If

    If Request.QueryString("Action") = "Do" Then
	'parse pasted address
	StrClip = Request.Form("PasteAddress")
    
	CurValue = GetField("Prefix", StrClip)
	If CurValue <> "" Then Prefix = CurValue
	
	CurValue = GetField("FirstName", StrClip)
	If CurValue <> "" Then FirstName = CurValue
	
	CurValue = GetField("LastName", StrClip)
	If CurValue <> "" Then LastName = CurValue
    
	CurValue = GetField("Firm", StrClip)
	If CurValue <> "" Then Company = CurValue
	
	CurValue = GetField("StreetAddress", StrClip)
	If CurValue <> "" Then StreetAddress = CurValue
	
	CurValue = GetField("StreetAddress2", StrClip)
	If CurValue <> "" Then StreetAddress2 = CurValue
	
	CurValue = GetField("City", StrClip)
	If CurValue <> "" Then City = CurValue
	
	CurValue = GetField("Province", StrClip)
	If CurValue <> "" Then Province = CurValue
	
	CurValue = GetField("Country", StrClip)
	If CurValue <> "" Then Country = CurValue
	
	CurValue = GetField("ZipCode", StrClip)
	If CurValue <> "" Then PostalCode = CurValue
	
	CurValue = GetField("Telephone", StrClip)
	If CurValue <> "" Then Telephone = CurValue
	
	CurValue = GetField("Fax", StrClip)
	If CurValue <> "" Then Fax = CurValue
	
	CurValue = GetField("Email", StrClip)
	If CurValue <> "" Then Email = CurValue

	CurValue = GetField("WebSite", StrClip)
	If CurValue <> "" Then WebSite = CurValue
	
	CurValue = GetField("Type", StrClip)
	If CurValue <> "" Then LicType = CurValue

	' now update the database
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
    
	ObjRec("Prefix") = Prefix 
	ObjRec("FirstName") = FirstName 
	ObjRec("LastName") = LastName 
	ObjRec("Company") = Company 
	ObjRec("StreetAddress") = StreetAddress 
	ObjRec("StreetAddress2") = StreetAddress2 
	ObjRec("City") = City  
	ObjRec("Province") = Province 
	ObjRec("PostalCode") = PostalCode 
	ObjRec("Email") = Email 
	ObjRec("Country") = Country 
	ObjRec("Telephone") = Telephone 
	ObjRec("Fax") = Fax 
	ObjRec("WebSite") = WebSite 
	ObjRec("Type") = LicType 

        objRec.Update
	
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
	
	ObjRec("ClientPrefix") = Prefix 
	ObjRec("ClientFirstName") = FirstName 
	ObjRec("ClientLastName") = LastName 
	ObjRec("ClientCompany") = Company 
	ObjRec("ClientStreetAddress") = StreetAddress 
	ObjRec("ClientStreetAddress2") = StreetAddress2 
	ObjRec("ClientCity") = City 
	ObjRec("ClientProvince") = Province 
	ObjRec("ClientCountry") = Country
	ObjRec("ClientPostalCode") = PostalCode 
	ObjRec("ClientEmail") = Email 
	ObjRec("ClientTelephone") = Telephone
	ObjRec("ClientFax") = Fax 

        objRec.Update
	
	'close database connection	
	objRec.Close
	Set objRec = Nothing
	
    Case Else
	nextPage = "/chi/cart/home.asp"
	
End Select	
	
	' redirect to the next page
	Select Case argType	
	    Case "License"
		nextPage = "/chi/licenses/license.asp?License=" & ID
	    Case "FranceLicense"
		nextPage = "/chi/france/license.asp?License=" & ID
	    Case "EnterpriseLicense"
		nextPage = "/chi/licenses/enterpriselicense.asp?EnterpriseLicense=" & ID
	    Case "FranceCompany"
		nextPage = "/chi/france/companylicense.asp?FranceLicense=" & ID
	    Case "Transaction"
		nextPage = "/chi/orders/order.asp?Transaction=" & ID
	End Select	

        Response.Redirect nextpage
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
<title>Paste Address for <% Response.Write argType & ": " & ID %> - Computational Hydraulics Int.</title>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache"> <% 'force page refresh %>
<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<META HTTP-EQUIV="EXPIRES" CONTENT="0">
<!meta description="Consulting engineering firm specializing in urban stormwater quantity and quality management modeling esp SWMM">
<!meta keywords="SWMM USEPA SWMM4.4 PCSWMM stormwater management model modeling modelling BMP's catchbasin civil engineering channel computer conduit discharge hydraulics hydrology management modeling modelling NPDES non-pointsource pipe pollutants precipitation programming quality rain rainfall runoff sediment simulation storm stormwater streamflow transport urban water quality water resources engineering">
</head>

<body text="#000000" bgcolor="#FFFFFF" link="#0000EE" vlink="#551A8B" alink="#FF0000" background="../cart/images/background.gif">

<p>
<form method="POST" action="paste.asp?Type=<%=argType%>&ID=<%=ID%>&Action=Do" name="FrontPage_Form1">
<table border="0" width="641">
  <tr>
    <td width="556">
        <p></p><font face="Arial" size="2"><b>Paste address from CHI database:</b><p>
        <textarea rows="12" name="PasteAddress" cols="66"></textarea></font></td>
    <td width="75" align="right" valign="top">
        <font face="Arial" size="2">
		<input type="submit" value="Submit" name="B1" style="float: right"></font></td>
  </tr>
  <tr>
    <td width="556">
        &nbsp;</td>
    <td width="75">
        &nbsp;</td>
  </tr>
</table>
</form>

</body>

</html>

<%

Function GetField(FieldName, StrClip)
    
    Dim Pos
    Dim RetPos

    GetField = ""
    
    Pos = InStr(1, StrClip, FieldName & "=")
    If Pos = 0 Then Pos = InStr(1, StrClip, FieldName & " =")
    If Pos = 0 Then Pos = InStr(1, StrClip, FieldName & vbTab & "=")
    If Pos > 0 Then
        Pos = InStr(Pos, StrClip, "=")
        RetPos = InStr(Pos, StrClip, vbCrLf)
        If RetPos = 0 Then RetPos = Len(StrClip) + 1
        GetField = Trim(Mid(StrClip, Pos + 1, RetPos - Pos - 1))
    End If

End Function

%>


















 
 