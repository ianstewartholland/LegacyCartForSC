<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<!--#include file="LogFile.asp"-->
<%
	'check to ensure CartItems exist
	If Session("NumCartItems") > 0 Then
		'everything ok
		If Session("SubTotalCost") <= 0 Or Session("PayQuoteMode") = "Yes" Then Response.Redirect("checkout.asp")
	Else
		Response.Redirect("cart.asp")
	End If

        ' session check to see if billing is same as client
        isDiffAsClient = False
        If Session("BillPrefix") <> Session("ClientPrefix") Or Session("BillFirstName") <> Session("ClientFirstName") Or Session("BillLastName") <> Session("ClientLastName") Or Session("BillName") <> Session("ClientName") Or Session("BillCompany") <> Session("ClientCompany") Or Session("BillStreetAddress") <> Session("ClientStreetAddress") Or Session("BillStreetAddress2") <> Session("ClientStreetAddress2") Or Session("BillCity") <> Session("ClientCity") Or Session("BillProvince") <> Session("ClientProvince") Or Session("BillCountry") <> Session("ClientCountry") Or Session("BillPostalCode") <> Session("ClientPostalCode") Or Session("BillEmail") <> Session("ClientEmail") Or Session("BillTelephone") <> Session("ClientTelephone") Or Session("BillFax") <> Session("ClientFax") Then isDiffAsClient = True
        
	If Request.QueryString("Action") = "UseClient" Then
            'swap sessions
            Session("BillPrefix") = Session("ClientPrefix")
            Session("BillFirstName") = Session("ClientFirstName")
            Session("BillLastName") = Session("ClientLastName")
            Session("BillName") = Session("ClientName")
            Session("BillCompany") = Session("ClientCompany")
            Session("BillStreetAddress") = Session("ClientStreetAddress")
            Session("BillStreetAddress2") = Session("ClientStreetAddress2")
            Session("BillCity") = Session("ClientCity")
            Session("BillProvince") = Session("ClientProvince")
            Session("BillCountry") = Session("ClientCountry")
            Session("BillPostalCode") = Session("ClientPostalCode")
            Session("BillEmail") = Session("ClientEmail")
            Session("BillTelephone") = Session("ClientTelephone")
            Session("BillFax") = Session("ClientFax")
            
            Response.Redirect "billingaddress.asp"
        End If

	If Request.QueryString("Action") = "Do" Then
	    Select Case Request.Form("Prefix")
		    Case "Please select (optional)..."
			    Session("BillPrefix") = ""
		    Case Else
			    Session("BillPrefix") = Request.Form("Prefix")
	    End Select
	    Session("BillFirstName") = Trim(Request.Form("FirstName"))
	    Session("BillLastName") = Trim(Request.Form("LastName"))
	    Session("BillName") = Trim(Session("BillPrefix") & " " & Trim(Session("BillFirstName") & " " & Session("BillLastName")))
	    Session("BillCompany") = Trim(Request.Form("Company"))
	    Session("BillStreetAddress") = Trim(Request.Form("StreetAddress"))
	    Session("BillStreetAddress2") = Trim(Request.Form("StreetAddress2"))
	    Session("BillCity") = Trim(Request.Form("City"))
            Session("BillCountry") = Trim(Request.Form("Country"))
	    Session("BillProvince") = Trim(Request.Form("Province"))
	    Session("BillPostalCode") = Trim(Request.Form("PostalCode"))
	    Session("BillEmail") = Trim(Request.Form("Email"))
	    Session("BillTelephone") = Trim(Request.Form("Telephone"))
	    Session("BillFax") = Trim(Request.Form("Fax"))

	    Session("BillAddress") = "Yes"
	    
	    ' server validation
	    If Session("BillProvince") = "" Then Response.Redirect "billingaddress.asp?Error=FillProvince"
	    If Session("BillFirstName") = "" Or Session("BillLastName") = "" Or Session("BillStreetAddress") = "" Or Session("BillCity") = "" Or Session("BillCountry") = "Please select a country..." Or Session("BillPostalCode") = "" Or Session("BillEmail") = "" Or Session("BillTelephone") = "" Then Response.Redirect "billingaddress.asp?Error=FillOut"
	    
	    If Session("BillCountry") = "USA" Or Session("BillCountry") = "Canada" Then
		outs = ValidateProvince (Session("BillCountry"), Session("BillProvince"))
		If Not outs(2) Then Response.Redirect "billingaddress.asp?Error=Province"
	
		Session("BillProvince") = outs(1)
	    End If
	    
	    If Not ValidateEmail(Session("BillEmail")) Then Response.Redirect "billingaddress.asp?Error=EmailAddr"
	    
            Session("ShowBilling") = True
            
	    Response.Redirect "checkout.asp"
	
	End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Billing Information - CHI Shopping Cart</title>
    <link rel="shortcut icon" href="assets/chilogo.ico" />
    <link href='https://fonts.googleapis.com/css?family=PT+Serif|Fjalla+One|Open+Sans:300italic,400italic,600italic,300,400,600,700'
        rel='stylesheet' type='text/css'>
    <link href="assets/site.min.css" rel="stylesheet" />
    <link href="assets/style.min.css" rel="stylesheet" />
    <script src="assets/startup.service.js"></script>
    <script src="assets/cart.js"></script>
</head>
<body>
    <!--#include file ="assets/header.asp"-->
    <div id="bodycontent" class="cart">
        <div class="container-fluid">
            <div class="container adjust-padding">
                <div class="row">
                    <div class="col-sm-12">
                        <h1>Billing Address</h1>
                        <%                        
        If Len(Request.QueryString("Error")) > 1 Then
	    Select Case Request.QueryString("Error")
		Case "Province"
		    errMsg = Session("BillProvince")
		    If Session("BillProvince") = ".." Then errMsg = "Please select..."
		    Response.Write "<p style='color:#FF0000; font-size:13px; font-weight:bold; margin:0px; margin-top:5px; padding-left:10px; padding-right:10px;'>The state/province '" & errMsg & "' you entered is not valid. Please try again.</p>"
		Case "EmailAddr"
		    Response.Write "<p style='color:#FF0000; font-size:13px; font-weight:bold; margin:0px; margin-top:5px; padding-left:10px; padding-right:10px;'>The email address '" & Session("BillEmail") & "' you entered is not valid. Please try again.</p>"
		Case "FillOut"
		    Response.Write "<p style='color:#FF0000; font-size:13px; font-weight:bold; margin:0px; margin-top:5px; padding-left:10px; padding-right:10px;'>Please fill out the required fields marked in red before proceeding.</p>"
		Case "FillProvince"
		    Response.Write "<p style='color:#FF0000; font-size:13px; font-weight:bold; margin:0px; margin-top:5px; padding-left:10px; padding-right:10px;'>Please enter a province, state, region or district before proceeding.</p>"  
	    End Select
        End If
                        %>
                        <p>
                            Edit your billing address here (if applicable).
                        </p>
                    </div>
                </div>
            </div>
        </div>

        <form method="POST" action="billingaddress.asp?Action=Do" name="BillingAddressForm" id="BillingAddressForm" onsubmit="return validate_form(this)"
            autocomplete="off">
            <div class="container-fluid white">
                <div class="container checkout">
                    <div class="row">
                        <div class="col-sm-12">
                            <div class="title-line">
                                Billing Information<% If isDiffAsClient Then Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href = 'billingaddress.asp?Action=UseClient' style='font-size: 12px;'>(Use client information)</a>" %>
                            </div>
                            <div class="table-row">
                                <div class="row">
                                    <div class="col-sm-3">Prefix</div>
                                    <div class="col-sm-9">
                                        
			<select size="1" name="Prefix">
                <option <% If Session("BillPrefix") = "" Then Response.Write "selected" %>>Please select (optional)...</option>
                <option <% If Session("BillPrefix") = "Mr." Then Response.Write "selected" %>>Mr.</option>
                <option <% If Session("BillPrefix") = "Ms." Then Response.Write "selected" %>>Ms.</option> 
                <option <% If Session("BillPrefix") = "Dr." Then Response.Write "selected" %>>Dr.</option>
                <option <% If Session("BillPrefix") = "Prof." Then Response.Write "selected" %>>Prof.</option>
            </select>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">First Name *</div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="FirstName" size="20" value="<% Response.Write Session("BillFirstName") %>" maxlength="20" <% If Session("BillAddress") = "Yes" And Session("BillFirstName") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Last Name *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="LastName" size="20" value="<% Response.Write Session("BillLastName") %>" maxlength="20" <% If Session("BillAddress") = "Yes" And Session("BillLastName") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Company/Organization *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="Company" size="40" value="<% Response.Write Session("BillCompany") %>" maxlength="50">
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Street address 1 *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="StreetAddress" size="40" value="<% Response.Write Session("BillStreetAddress") %>" maxlength="40"
          <% If Session("BillAddress") = "Yes" And Session("BillStreetAddress") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Street address 2 (optional)
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="StreetAddress2" size="40" value="<% Response.Write Session("BillStreetAddress2") %>" maxlength="40">
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        City *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="City" size="20" value="<% Response.Write Session("BillCity") %>" maxlength="20" <% If Session("BillAddress") = "Yes" And Session("BillCity") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <%
If Not Session("ShowBilling") Then
                                %>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Country *
                                    </div>
                                    <div class="col-sm-9">
                                        
        <select name="Country" id="Country" onchange='setStates();'>
            <option>Please select a country...</option>
            <option <% If Session("BillCountry") = "USA" Then Response.Write "selected" %>>USA</option>
            <option <% If Session("BillCountry") = "Canada" Then Response.Write "selected" %>>Canada</option>

            <%
        Set objRec = Server.CreateObject ("ADODB.Recordset")
    	objRec.Open "Select Country From ShippingCalculator ORDER BY Country ASC", shippingcalculatorPath
        
        If Not objRec.EOF then
            Do While Not objRec.EOF    
		If objRec("Country") <> "Canada" And objRec("Country") <> "USA" Then
		    If Session("BillCountry") = objRec("Country") Then
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
                                    </div>
                                </div>
                                <% Else %>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Country
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="Country" style="background-color: #FFF; border: none;" size="20" value="<%= Session("BillCountry") %>"
          readonly>
                                    </div>
                                </div>
                                <% End If %>

                                <% If Session("BillAddress") <> "Yes" And Session("BillCountry") <> "USA" And Session("BillCountry") <> "Canada" Then %>

                                <div class="row">
                                    <div class="col-sm-3">
                                        Province/State *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="Province" id="Province" size="20" value="<%=Session("BillProvince") %>" <% If Session("BillAddress") = "Yes" And Session("BillProvince") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>

                                <%
Else
DisplayProvinceDropdown Session("BillCountry"), Session("BillProvince")

End If
                                %>

                                <div class="row">
                                    <div class="col-sm-3">

                                        <%
                Select Case Session("BillCountry")
                    Case "USA"
                        Response.Write "Zip code"
		    Case Else
                        Response.Write "Postal code"
                End select
                                        %>
             *
                                        
                                    </div>
                                    <div class="col-sm-9">
                                        
	    <input type="text" name="PostalCode" size="20" value="<% Response.Write Session("BillPostalCode") %>" maxlength="15" <% If Session("BillAddress") = "Yes" And Session("BillPostalCode") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Email *
                                    </div>
                                <div class="col-sm-9">
                                    
	    <input type="text" name="Email" size="60" value="<% Response.Write Session("BillEmail") %>" maxlength="65" <% If Session("BillAddress") = "Yes" And Session("BillEmail") = "" Then Response.Write "class='txtBackground'" %>>
                                </div>
                            </div>
                            <div class="row">
                                <div class="col-sm-3">
                                    Telephone *<br>
                                    (e.g. 123-123-1234)
                                </div>
                                <div class="col-sm-9">
                                    
	  <input type="text" name="Telephone" size="20" value="<% Response.Write Session("BillTelephone") %>" maxlength="25" <% If Session("BillAddress") = "Yes" And Session("BillTelephone") = "" Then Response.Write "class='txtBackground'" %>>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <br class="hidden-sm hidden-xs" />
    </div>

    <div class="container-fluid">
        <div class="container">
            <div class="row">
                <div class="col-sm-12">
                    <div class="row button-row">
                        <div class="col-xs-12 right">
                            <button type="submit" form="BillingAddressForm" class="btn btn-blue">
                                UPDATE
                                    <img class="arrow" src="assets/arrow-right.png" /></button>

                            <a href="checkout.asp" class="btn btn-blue">CANCEL
                                    <img class="arrow" src="assets/arrow-right.png" /></a>
                        </div>
                    </div>
                    <div class="more-info">
                        <p class="title">NEED MORE INFORMATION?</p>
                        <p>
                            If you have any questions, please email <a href="mailto:info@chiwater.com">info@chiwater.com</a> or call us at 1-519-767-0197
                            or 1-888-972-7966.
                        </p>
                    </div>
                </div>
            </div>
        </div>
    </div>
    </form>

    </div>

    <!--#include file ="assets/footer.asp"-->
</body>
</html>
