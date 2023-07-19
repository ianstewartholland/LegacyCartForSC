<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<!--#include file="LogFile.asp"-->
<%
	'check to ensure CartItems exist
	If Session("NumCartItems") > 0 Then
		'everything ok
		If Session("SubTotalCost") <= 0 Or Session("PayQuoteMode") = "Yes" Or Session("ShippingRequired") <> "Yes" Then Response.Redirect("checkout.asp")
	Else
		Response.Redirect("cart.asp")
	End If

        ' session check to see if shipping is same as client
        isDiffAsClient = False
        If Session("ShipPrefix") <> Session("ClientPrefix") Or Session("ShipFirstName") <> Session("ClientFirstName") Or Session("ShipLastName") <> Session("ClientLastName") Or Session("ShipName") <> Session("ClientName") Or Session("ShipCompany") <> Session("ClientCompany") Or Session("ShipStreetAddress") <> Session("ClientStreetAddress") Or Session("ShipStreetAddress2") <> Session("ClientStreetAddress2") Or Session("ShipCity") <> Session("ClientCity") Or Session("ShipProvince") <> Session("ClientProvince") Or Session("ShipCountry") <> Session("ClientCountry") Or Session("ShipPostalCode") <> Session("ClientPostalCode") Or Session("ShipEmail") <> Session("ClientEmail") Or Session("ShipTelephone") <> Session("ClientTelephone") Or Session("ShipFax") <> Session("ClientFax") Then isDiffAsClient = True
 
	If Request.QueryString("Action") = "UseClient" Then
            'swap sessions
            Session("ShipPrefix") = Session("ClientPrefix")
            Session("ShipFirstName") = Session("ClientFirstName")
            Session("ShipLastName") = Session("ClientLastName")
            Session("ShipName") = Session("ClientName")
            Session("ShipCompany") = Session("ClientCompany")
            Session("ShipStreetAddress") = Session("ClientStreetAddress")
            Session("ShipStreetAddress2") = Session("ClientStreetAddress2")
            Session("ShipCity") = Session("ClientCity")
            Session("ShipProvince") = Session("ClientProvince")
            Session("ShipCountry") = Session("ClientCountry")
            Session("ShipPostalCode") = Session("ClientPostalCode")
            Session("ShipEmail") = Session("ClientEmail")
            Session("ShipTelephone") = Session("ClientTelephone")
            Session("ShipFax") = Session("ClientFax")
            
            Response.Redirect "shippingaddress.asp"
        End If
        
	If Request.QueryString("Action") = "Do" Then
	    '
	    ' this page may seem wired but it's to prevent when client side's validation fails, people try to enter invalid province/country
	    '
	    Select Case Request.Form("Prefix")
		    Case "Please select (optional)..."
			    ShipPrefix = ""
		    Case Else
			    ShipPrefix = Request.Form("Prefix")
	    End Select
	    ShipFirstName = Trim(Request.Form("FirstName"))
	    ShipLastName = Trim(Request.Form("LastName"))
	    ShipName = Trim(ShipPrefix) & " " & Trim(ShipFirstName & " " & ShipLastName)
	    ShipCompany = Trim(Request.Form("Company"))
	    ShipStreetAddress = Trim(Request.Form("StreetAddress"))
	    ShipStreetAddress2 = Trim(Request.Form("StreetAddress2"))

	    ShipCity = Trim(Request.Form("City"))
	    ShipCountry = Trim(Request.Form("Country"))
	    ShipProvince = Trim(Request.Form("Province"))
			       
	    ShipPostalCode = Trim(Request.Form("PostalCode"))
	    ShipEmail = Trim(Request.Form("Email"))
	    ShipTelephone = Trim(Request.Form("Telephone"))
	    ShipFax = Trim(Request.Form("Fax"))
	    
	    Session("ShipAddress") = "Yes"
	    
	    ' validation
	    If ShipProvince = "" Then Response.Redirect "shippingaddress.asp?Error=FillProvince"
	    If ShipFirstName = "" Or ShipLastName = "" Or ShipStreetAddress = "" Or ShipCity = "" Or ShipCountry = "Please select a country..." Or ShipPostalCode = "" Or ShipEmail = "" Or ShipTelephone = "" Then Response.Redirect "shippingaddress.asp?Error=FillOut"
	   
	    If ShipCountry = "USA" Or ShipCountry = "Canada" Then
		outs = ValidateProvince (ShipCountry, ShipProvince)
		If Not outs(2) Then Response.Redirect "shippingaddress.asp?Error=Province"

		ShipProvince = outs(1)
	    End If
	    
	    If Not ValidateEmail(ShipEmail) Then Response.Redirect "shippingaddress.asp?Error=EmailAddr"

	    If Session("ShipCountry") <> Trim(Request.Form("Country")) Or Session("ShipProvince") <> Trim(Request.Form("Province")) Or Session("ShipCity") <> Trim(Request.Form("City")) Or Session("ShipPostalCode") <> Trim(Request.Form("PostalCode")) Then
		' reset currency
		If Session("ShipCountry") <> Trim(Request.Form("Country")) Then
		    If Trim(Request.Form("Country")) = "Canada" Then
			Session("Currency") = "Canadian"
			Session("CurrencyAbrev") = "CAD"
		    ElseIf Trim(Request.Form("Country")) = "USA" Then
			Session("Currency") = "USA"
			Session("CurrencyAbrev") = "US"
		    Else
			Session("Currency") = "International"
			Session("CurrencyAbrev") = "US"
		    End If
		End If
	    End If
	    
	    ' now update sessions
            Session("ShipPrefix") = ShipPrefix
            Session("ShipFirstName") = ShipFirstName
            Session("ShipLastName") = ShipLastName
            Session("ShipName") = ShipName
            Session("ShipCompany") = ShipCompany
            Session("ShipStreetAddress") = ShipStreetAddress
            Session("ShipStreetAddress2") = ShipStreetAddress2
            Session("ShipCity") = ShipCity
            Session("ShipProvince") = ShipProvince
            Session("ShipCountry") = ShipCountry
            Session("ShipPostalCode") = ShipPostalCode
            Session("ShipEmail") = ShipEmail
            Session("ShipTelephone") = ShipTelephone
            Session("ShipFax") = ShipFax

	    Session("ShowShipping") = True
	    'check if a client address exists, if not then assign shipping address as default
	    ' old code, I know we might never need this but just leave it here for reference
	    If Session("ClientAddress") <> "Yes" Then
		    Session("ClientPrefix") = Session("ShipPrefix")
		    Session("ClientFirstName") = Session("ShipFirstName")
		    Session("ClientLastName") = Session("ShipLastName")
		    Session("ClientName") = Session("ShipName")
		    Session("ClientCompany") = Session("ShipCompany")
		    Session("ClientStreetAddress") = Session("ShipStreetAddress")
		    Session("ClientStreetAddress2") = Session("ShipStreetAddress2")
		    Session("ClientCity") = Session("ShipCity")
		    Session("ClientProvince") = Session("ShipProvince")
		    Session("ClientCountry") = Session("ShipCountry")
		    Session("ClientPostalCode") = Session("ShipPostalCode")
		    Session("ClientEmail") = Session("ShipEmail")
		    Session("ClientTelephone") = Session("ShipTelephone")
		    Session("ClientFax") = Session("ShipFax")
	    End If
	    'check if a billing address exists, if not then assign shipping address as default
	    If Session("BillAddress") <> "Yes" Then
		    Session("BillPrefix") = Session("ShipPrefix")
		    Session("BillFirstName") = Session("ShipFirstName")
		    Session("BillLastName") = Session("ShipLastName")
		    Session("BillName") = Session("ShipName")
		    Session("BillCompany") = Session("ShipCompany")
		    Session("BillStreetAddress") = Session("ShipStreetAddress")
		    Session("BillStreetAddress2") = Session("ShipStreetAddress2")
		    Session("BillCity") = Session("ShipCity")
		    Session("BillProvince") = Session("ShipProvince")
		    Session("BillCountry") = Session("ShipCountry")
		    Session("BillPostalCode") = Session("ShipPostalCode")
		    Session("BillEmail") = Session("ShipEmail")
		    Session("BillTelephone") = Session("ShipTelephone")
		    Session("BillFax") = Session("ShipFax")
	    End If
    
	    Response.Redirect "checkout.asp"

	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- #BeginTemplate "/Templates/innerpage2.dwt" --><!-- DW6 -->

<!-- #BeginEditable "variables" -->
<% 
Dim mainpage, subpage
mainpage = ""
subpage = "Shipping Address"
%>
<!-- #EndEditable -->

<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<link href="styles/style2.css" rel="STYLESHEET" type="text/css" media="screen" />
	<link href="styles/print.css" rel="STYLESHEET" type="text/css" media="print" />
	<link rel="shortcut icon" href="images/chilogo.ico" />
	<!-- #BeginEditable "title" -->
<title>Shipping Address - CHI</title>
<script type="text/javascript" src="styles/cart.js"></script>
<!-- #EndEditable -->
</head>
<body <% If Len(Session("ShipProvince")) < 1 Then Response.Write "onload='pageLoad();'" %>>
<div class="main">
		<div class="logo_home">			
			<div class="header">
				<a style="width:150px;height:100px;float:left;text-decoration:none;" href="http://www.chiwater.com">&nbsp;</a>
				<div class="slogan">
					<!-- #BeginEditable "slogan" -->
					    Dedicated to promoting healthier ecosystems and communities<br />
					    by providing the software, resources and guidance to simplify<br />
					    advanced water management and design.
					<!-- #EndEditable -->
				</div>
<div class="print_search">
	<form action="http://www.chiwater.com/search_results.asp" id="searchform">
<%
If Session("NumCartItems") <> "" Then
    If Session("NumCartItems") > 1 Then
        Response.Write "<a style='color: #FFFFFF; link: #FFFFFF; visited: #FFFFFF; font-size:11px; font-weight:bold;' class='nav' href = 'https://secure.chiwater.com/chi/cart/cart.asp'>Cart (" & Session("NumCartItems") & " items)</a>&nbsp;&nbsp;"
    ElseIf Session("NumCartItems") = 1 Then
        Response.Write "<a style='color: #FFFFFF; link: #FFFFFF; visited: #FFFFFF; font-size:11px; font-weight:bold;' class='nav' href = 'https://secure.chiwater.com/chi/cart/cart.asp'>" &  "Cart (1 item)</a>&nbsp;&nbsp;"
    End If
End If
%>
		<input type="text" name="q" value="Search..." class="searchbox" onfocus="javascript:this.value='';" onblur="javascript:this.value='Search...';this.style.color = '#808080';"/>
		<input type="image" src="images/search_arrow.png" style="vertical-align:top;"/>
	</form>
</div>
			</div>			
			
			<div style="clear:both;"></div>
		</div>
		<div class="banner"></div>
		<div class="innerbg">
			<ul>

				<li class="main_content_large">
				
				<!-- #BeginEditable "Body" -->
					<ul>
						<li class="maincontenttop_large"></li>
						<li class="maincontentarea_large">	
            
			<h1>Shipping Address</h1>
<%                        
        If Len(Request.QueryString("Error")) > 1 Then
	    Select Case Request.QueryString("Error")
		Case "Province"
		    errMsg = Session("ShipProvince")
		    If Session("ShipProvince") = ".." Then errMsg = "Please select..."
		    Response.Write "<p style='color:#FF0000; font-size:13px; font-weight:bold; margin:0px; margin-top:5px; padding-left:10px; padding-right:10px;'>The state/province '" & errMsg & "' you entered is not valid. Please try again.</p>"
		Case "EmailAddr"
		    Response.Write "<p style='color:#FF0000; font-size:13px; font-weight:bold; margin:0px; margin-top:5px; padding-left:10px; padding-right:10px;'>The email address '" & Session("ShipEmail") & "' you entered is not valid. Please try again.</p>"
		Case "FillOut"
		    Response.Write "<p style='color:#FF0000; font-size:13px; font-weight:bold; margin:0px; margin-top:5px; padding-left:10px; padding-right:10px;'>Please fill out the required fields marked in red and check country/province before proceeding.</p>"
		Case "FillProvince"
		    Response.Write "<p style='color:#FF0000; font-size:13px; font-weight:bold; margin:0px; margin-top:5px; padding-left:10px; padding-right:10px;'>Please enter a province, state, region or district before proceeding.</p>"  
	    End Select
        End If
%>  
			<p>Edit your shipping or mailing address here (if applicable).</p>
<form method="POST" action="shippingaddress.asp?Action=Do" name="ShippingAddressForm" onsubmit="return validate_form(this)" autocomplete = "off">			
<table width="96%">
    <tr>
        <td colspan = '2' bgcolor="#B3C6CA">
	    <h3>Shipping Information<% If isDiffAsClient Then Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href = 'shippingaddress.asp?Action=UseClient'>Use client information</a>" %>
            </h3></td>
    </tr>
        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Prefix</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<select size="1" name="Prefix">
			<option <% If Session("ShipPrefix") = "" Then Response.Write "selected" %>>Please select (optional)...</option>
			<option <% If Session("ShipPrefix") = "Mr." Then Response.Write "selected" %>>Mr.</option>
			<option <% If Session("ShipPrefix") = "Ms." Then Response.Write "selected" %>>Ms.</option>
			<option <% If Session("ShipPrefix") = "Dr." Then Response.Write "selected" %>>Dr.</option>
			<option <% If Session("ShipPrefix") = "Prof." Then Response.Write "selected" %>>Prof.</option>
			</select></td>
        </tr>
		<tr>
          <td width="30%" bgcolor="#F4F8F9"><p>First Name *</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<input type="text" name="FirstName" size="20" value="<% Response.Write Session("ShipFirstName") %>" maxlength="20" <% If Session("ShipAddress") = "Yes" And Session("ShipFirstName") = "" Then Response.Write "class='txtBackground'" %>></td>
        </tr>
        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Last Name *</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<input type="text" name="LastName" size="20" value="<% Response.Write Session("ShipLastName") %>" maxlength="20" <% If Session("ShipAddress") = "Yes" And Session("ShipLastName") = "" Then Response.Write "class='txtBackground'" %>></td>
        </tr>
        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Company/Organization *</p>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
                        <input type="text" name="Company" size="40" value="<% Response.Write Session("ShipCompany") %>" maxlength="50"></td>
        </tr>
        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Street address 1 *</p>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<input type="text" name="StreetAddress" size="40" value="<% Response.Write Session("ShipStreetAddress") %>" maxlength="40" <% If Session("ShipAddress") = "Yes" And Session("ShipStreetAddress") = "" Then Response.Write "class='txtBackground'" %>></td>
        </tr>
        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Street address 2 (optional)</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<input type="text" name="StreetAddress2" size="40" value="<% Response.Write Session("ShipStreetAddress2") %>" maxlength="40"></td>
        </tr>
        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>City *</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<input type="text" name="City" size="20" value="<% Response.Write Session("ShipCity") %>" maxlength="20" <% If Session("ShipAddress") = "Yes" And Session("ShipCity") = "" Then Response.Write "class='txtBackground'" %>></td>
        </tr>
<%
If Not Session("ShowShipping") Then
%>
        <tr>
            <td width="30%" bgcolor="#F4F8F9"><p>Country *</p></td>
            <td width="70%" bgcolor="#F4F8F9">&nbsp;
        <select name="Country" id="Country" onchange='setStates();'>
        <option>Please select a country...</option>
        <option <% If Session("ShipCountry") = "USA" Then Response.Write "selected" %>>USA</option>
        <option <% If Session("ShipCountry") = "Canada" Then Response.Write "selected" %>>Canada</option>
        
        <%
        Set objRec = Server.CreateObject ("ADODB.Recordset")
    	'strConnect = "Driver={Microsoft Access Driver (*.mdb)};DBQ=C:\WWWUsers\LocalUser\chi\computationalhydraulics_com\databases\ShippingCalculator2010.mdb"
    	objRec.Open "Select Country From ShippingCalculator ORDER BY Country ASC", shippingcalculatorPath
        
        If Not objRec.EOF then
            Do While Not objRec.EOF    
		If objRec("Country") <> "Canada" And objRec("Country") <> "USA" Then
		    If Session("ShipCountry") = objRec("Country") Then
			Response.Write "<option selected>" & objRec("Country") & "</option>"
		    Else
			Response.Write "<option>" & objRec("Country") & "</option>"
		    End If
		    objRec.MoveNext
		Else
		    objRec.MoveNext
		End If
            Loop
        objRec.Close
        End If
        %>
        
        </select> 
            </td>
        </tr>

<% Else %>
    <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Country</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
	  <input type="text" name="Country" size="20" value="<%= Session("ShipCountry") %>" readonly style="color:#314B5B;font-family:Arial, Helvetica, sans-serif;font-size:12px;background-color:#F4F8F9;border: none;"></td>
    </tr> 
<% End If %>	
<% If Session("ShipAddress") <> "Yes" And Session("ShipCountry") <> "USA" And Session("ShipCountry") <> "Canada" Then %>
    
    <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Province/State *</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
	  <input type="text" name="Province" id="Province" size="20" value="<%=Session("ShipProvince") %>" <% If Session("ShipAddress") = "Yes" And Session("ShipProvince") = "" Then Response.Write "class='txtBackground'" %>></td>
    </tr>
 
<%
Else
DisplayProvinceDropdown Session("ShipCountry"), Session("ShipProvince")

End If
%>

        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>
<%
                Select Case Session("ShipCountry")
                    Case "USA"      
                        Response.Write "Zip code"          
                    Case Else
                        Response.Write "Postal code"
                End select
%>
           *</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<input type="text" name="PostalCode" size="20" value="<% Response.Write Session("ShipPostalCode") %>" maxlength="15" <% If Session("ShipAddress") = "Yes" And Session("ShipPostalCode") = "" Then Response.Write "class='txtBackground'" %>></td>
        </tr>
        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Email *</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<input type="text" name="Email" size="60" value="<% Response.Write Session("ShipEmail") %>" maxlength="65" <% If Session("ShipAddress") = "Yes" And Session("ShipEmail") = "" Then Response.Write "class='txtBackground'" %>></td>
        </tr>
        <tr>
          <td width="30%" bgcolor="#F4F8F9"><p>Telephone *</p></td>
          <td width="70%" bgcolor="#F4F8F9">&nbsp;
			<input type="text" name="Telephone" size="20" value="<% Response.Write Session("ShipTelephone") %>" maxlength="25" <% If Session("ShipAddress") = "Yes" And Session("ShipTelephone") = "" Then Response.Write "class='txtBackground'" %>></td>
        </tr>
      </table>

<table width="96%">
    <tr>
	<td align="right"><br>

	    <input type="image" src="images/button_update.png" value="ClientInfo" name="I1">&nbsp;
	    <a href="checkout.asp"><img border="0" src="images/button_cancel.png"></a>

	</td>
    </tr>
</table>			
			
    </form>			
			

                      <li class="maincontentbottom_large"></li>
</ul>	                   
				               
					<p class="p_no_margin">&nbsp;</p>
        <ul>
	<li class="maincontenttop2_large">
<h4>
For more information</h4></li>
	<li class="maincontentarea_large">
<p>If you have any questions, please email <a href="mailto:info@chiwater.com">info@chiwater.com</a>,
    or call us at telephone: 1-519-767-0197 or 1-888-972-7966. </p>
      </li>
	<li class="maincontentbottom_large"></li>
</ul>    
															
				<!-- #EndEditable -->
				</li>
				<li class="between">&nbsp;</li>
				<li class="right_content">
				
					<!-- #BeginEditable "RightBar" -->
<!-- #include file = "styles/leftmenu_cart.asp" -->
					<!-- #EndEditable -->	
			  </li>
				<li class="start_end">&nbsp;</li>
		  </ul>
</div>
		
		<div class="footer">
<div style="float:left">
	Copyright © <%= Year(Date) %> CHI. All rights reserved.
</div>

<div style="float:right;text-align:right">	
	<a href="http://www.chiwater.com/Disclaimer.asp">Disclaimer</a>
	 | 
	<a href="http://www.chiwater.com/PrivacyPolicy.asp">Privacy Policy</a>
</div>
		</div>
		
	</div>
</body>
<!-- #EndTemplate --></html>