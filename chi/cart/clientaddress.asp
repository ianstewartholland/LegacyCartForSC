﻿<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<!--#include file="LogFile.asp"-->
<!--#include virtual ="/chi/includes/taxcalculator.asp"-->
<%
	'check to ensure CartItems exist
	If Session("NumCartItems") > 0 Then
		'everything ok
		If Session("PayQuoteMode") = "Yes" And Not IsEmpty(Session("ConferenceInterests")) Then Response.Redirect("checkout.asp")
	Else
		Response.Redirect("cart.asp")
	End If

	Session("AddressClient") = False
    
    If Len(Session("Consent")) = 0 Then Session("Consent") = -1

        If Session("NumCartItems") = 1 And Session("CartItem(1)") = "S212" Then
            isTrial = True
        Else
            isTrial = False
        End If

	' branch for google search question
	isConference = False
	isTraining = False
	isSoftware = False
	Question = ""

    isWebTraining = False
    isSpecialty = False

    DayAll = True
    DayOne = False
    DayTwo = False

	For t = 1 to Session("NumCartItems")
	    If Left(Session("CartItem(" & t & ")"),1) = "C" Then
		isConference = True
	    ElseIf Left(Session("CartItem(" & t & ")"),1) = "W" Then
		isTraining = True
            If InStr(Session("CartItemDescription(" & t & ")"), "South Africa") < 1 Then
                If Right(Session("CartItem(" & t & ")"),1) = "A" Then
                    DayOne = True
                ElseIf Right(Session("CartItem(" & t & ")"),1) = "B" Then
                    DayTwo = True
                End If
            End If
	    ElseIf Left(Session("CartItem(" & t & ")"),1) = "S" Or Left(Session("CartItem(" & t & ")"),1) = "P" Then
		isSoftware = True
	    End If
        
        If InStr(Session("CartItemDescription(" & t & ")"), WebTrainingItem) > 0 Then isWebTraining = True
        If InStr(Session("CartItemDescription(" & t & ")"), "specialty") > 0 Then isSpecialty = True
        
	Next

    If (DayOne Or DayTwo) Then DayAll = False

	If isConference Then
	    Question = "How did you find out about CHI's conference?"
	ElseIf isTraining Then
	    Question = "How did you find out about CHI's training?"
	ElseIf isSoftware Then
	    Question = "How did you find out about PCSWMM?"
	Else
	    Question = "How did you find out about CHI's products/services?"
	End If

	If Request.QueryString("Action") = "Do" Then

            Session("AddressClient") = True

	    Session("ShowClient") = False
            ' check email address for trial license
            If isTrial And EmailCheckFailed(Trim(Request.Form("Email"))) Then Response.Redirect "clientaddress.asp?Error=Email"
            
        ' workshop interests
        Session("WorkshopInterests") = ""
        WorkshopInterests = False
        If Request.Form("WInterest1") <> "Please select..." Then
            Session("WInterest1") = Request.Form("WInterest1") 
            WorkshopInterests = True
        Else
            Session("WInterest1") = 0
        End If
        If Request.Form("WInterest2") <> "Please select..." Then
            Session("WInterest2") = Request.Form("WInterest2") 
            WorkshopInterests = True
        Else
            Session("WInterest2") = 0
        End If
        If Request.Form("WInterest3") <> "Please select..." Then
            Session("WInterest3") = Request.Form("WInterest3") 
            WorkshopInterests = True
        Else
            Session("WInterest3") = 0
        End If
        
        If Request.Form("WExperience1") <> "Please select..." Then
            Session("WExperience1") = Request.Form("WExperience1") 
            WorkshopInterests = True
        Else
            Session("WExperience1") = 0
        End If

        If Request.Form("WExperience2") <> "Please select..." Then
            Session("WExperience2") = Request.Form("WExperience2") 
            WorkshopInterests = True
        Else
            Session("WExperience2") = 0
        End If

        If WorkshopInterests Then Session("WorkshopInterests") = "Yes"

        ' conference interests
        Session("ConferenceInterests") = ""
        If Request.Form("Interest1") <> "Please select..." Then
            Session("ConferenceInterests") = Session("ConferenceInterests") & Request.Form("Interest1") + "|"
        End If
        If Request.Form("Interest2") <> "Please select..." Then
            Session("ConferenceInterests") = Session("ConferenceInterests") & Request.Form("Interest2") + "|"
        End If
        If Request.Form("Interest3") <> "Please select..." Then
            Session("ConferenceInterests") = Session("ConferenceInterests") & Request.Form("Interest3") 
        End If
        
        If Session("ConferenceInterests") = "" Then Session("ConferenceInterests") = "No interest declared"

                Select Case Request.Form("Prefix")
                    Case "Please select (optional)..."
                            Session("ClientPrefix") = ""
                    Case Else
                            Session("ClientPrefix") = Request.Form("Prefix")
                End Select
                Session("ClientFirstName") = Trim(Request.Form("FirstName"))
                Session("ClientLastName") = Trim(Request.Form("LastName"))
                Session("ClientName") = Trim(Session("ClientPrefix") & " " & Trim(Session("ClientFirstName") & " " & Session("ClientLastName")))
                Session("ClientCompany") = Trim(Request.Form("Company"))
                Session("ClientStreetAddress") = Trim(Request.Form("StreetAddress"))
                Session("ClientStreetAddress2") = Trim(Request.Form("StreetAddress2"))
                Session("ClientCity") = Trim(Request.Form("City"))

                ' set currency on country selection
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

                Session("ClientCountry") = Trim(Request.Form("Country"))
		        Session("ClientProvince") = Trim(Request.Form("Province"))
                Session("ClientPostalCode") = Trim(Request.Form("PostalCode"))
                Session("ClientEmail") = Trim(Request.Form("Email"))
                Session("ClientTelephone") = Trim(Request.Form("Telephone"))
                Session("ClientFax") = Trim(Request.Form("Fax"))
                Session("ClientWebSite") = Trim(Request.Form("WebSite"))

                Select Case Trim(Request.Form("Type"))
                        Case "Municipality"
                                Session("ClientType") = "MA"
                        'Case "Regional"
                        '        Session("ClientType") = "RA"
                        Case "Government Agency"
                                Session("ClientType") = "GA"
                        Case "Organization/Association"
                                Session("ClientType") = "OR"
                        Case "Consulting Company"
                                Session("ClientType") = "CE"
                        Case "University/College"
                                Session("ClientType") = "UN"
                        Case "Student"
                                Session("ClientType") = "ST"
                        Case "Commercial/Industrial"
                                Session("ClientType") = "IN"
                        'Case "Library"
                        '        Session("ClientType") = "LB"
                        'Case "Book Dealer"
                        '        Session("ClientType") = "BK"
                End Select

		Session("SourceOption") = -1
		Select Case Request.Form("Source")
		    Case "Email from CHI"
			Session("SourceOption") = 1
		    Case "Web search"
			Session("SourceOption") = 2
		   'Case "By checking the CHI website"
			'Session("SourceOption") = 3
		   'Case "An ad on another website"
			'Session("SourceOption") = 4
		    Case "Colleague"
			Session("SourceOption") = 5
		   'Case "Paper flyer in the mail"
			'Session("SourceOption") = 6
		   'Case "Telephone call from CHI"
			'Session("SourceOption") = 7
            ' start collecting since mar 2020
		    Case "Anthropocene Engineering Sales Representative"
			Session("SourceOption") = 15
            ' start collecting since dec 2021
		    Case "Social Media"
			Session("SourceOption") = 16
		    Case "Please select..."
			Session("SourceOption") = 0
		End Select

        If Left(Request.Form("Consent"), 2) = "No" Then 
            Session("Consent") = 0
        Else 'If Left(Request.Form("Consent"), 2) = "Ye" Then 
            Session("Consent") = 1
        End If

		Session("ClientAddress") = "Yes"

	    '
	    ' server validation
	    '
	    If Session("ClientProvince") = "" Then Response.Redirect "clientaddress.asp?Error=FillProvince"
	    If Session("ClientFirstName") = "" Or Session("ClientLastName") = "" Or Session("ClientCity") = "" Or Session("ClientCountry") = "Please select a country..." Or Session("ClientPostalCode") = "" Or Session("ClientEmail") = "" Or Session("ClientTelephone") = "" Then Response.Redirect "clientaddress.asp?Error=FillOut"

	    If Session("ClientCountry") = "USA" Or Session("ClientCountry") = "Canada" Then
		outs = ValidateProvince (Session("ClientCountry"), Session("ClientProvince"))
		If Not outs(2) Then Response.Redirect "clientaddress.asp?Error=Province"

		Session("ClientProvince") = outs(1)
	    End If

	    If Not ValidateEmail(Session("ClientEmail")) Then Response.Redirect "clientaddress.asp?Error=EmailAddr"
                'check if a shipping address exists, if not then assign client address as default
                If Session("ShipAddress") <> "Yes" Or Session("ShippingRequired") = "Yes" Then
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
                End If
                'check if a billing address exists, if not then assign client address as default
                If Session("BillAddress") <> "Yes" Then
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
                End If
                Session("ShowClient") = True
                Session("ShowShipping") = False
                Session("ShowBilling") = False
            If Not isTrial Then
' to test
'    response.write Session("ConferenceInterests")
                Response.Redirect "checkout.asp"
            Else
                ' trial order
                Session("ShippingCost") = 0
		        Tax = TaxCalculator (Session("ShipProvince"), Session("CurrencyAbrev"), ItemStr, Session("ShippingCost"), "", "", "")

		        Session("GSTCost") = Tax(0)
		        Session("PSTCost") = Tax(1)
		        Session("HSTCost") = Tax(2)
		        Session("TotalCost") = Tax(3)
		        Session("SubTotalCost") = Tax(4)

                Response.Redirect "ordercommit.asp"
            End If
	End If

%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Client Information - CHI Shopping Cart</title>
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
                        <h1>Client Information</h1>
                        <%
        If Len(Request.QueryString("Error")) > 1 Then
	    Select Case Request.QueryString("Error")
		Case "Email"
		    Response.Write "<p class='error'>Unfortunately we do not accept personal email addresses for PCSWMM trial licenses. If you absolutely need to use this address, please call us at 519-767-0197 or email us at <a href = 'mailto:info@chiwater.com'>info@chiwater.com</a>.</p>"
		Case "Province"
		    errMsg = Session("ClientProvince")
		    If Session("ClientProvince") = ".." Then errMsg = "Please select..."
		    Response.Write "<p class='error'>The state/province '" & errMsg & "' you entered is not valid. Please try again.</p>"
		Case "EmailAddr"
		    Response.Write "<p class='error'>The email address '" & Session("ClientEmail") & "' you entered is not valid. Please try again.</p>"
		Case "FillOut"
		    Response.Write "<p class='error'>Please fill out the required fields marked in red before proceeding.</p>"
		Case "FillProvince"
		    Response.Write "<p class='error'>Please enter a province, state, region or district before proceeding. Enter n/a if not applicable.</p>"
	    End Select
        End If

        If Session("ClientQuote") Then
            Response.Write "<p><b>To view a quote showing applicable taxes and shipping (if any), your information is needed.</b></p>"
        End If
                        %>
                        <p>
                            Enter the name, address and contact information of the person who will be considered the client (e.g. attendee or license
                            holder, if applicable).
                        </p>
                        <% If isTrial Then %>
                        <p>
                            <b>Please note: If you are a full time student with a student email address and require PCSWMM for educational purposes please
                                visit our <a target="_blank" href="https://www.pcswmm.com/Grant">Grant</a> program. Students please use the University mailing
                                address on your request.</b>
                        </p>
                        <% End If %>
                    </div>
                </div>
            </div>
        </div>

        <form method="POST" action="clientaddress.asp?Action=Do" name="ClientAddressForm" id="ClientAddressForm" onsubmit="return validate_form(this)"
            autocomplete="off">
            <div class="container-fluid white">
                <div class="container checkout">
                    <div class="row">
                        <div class="col-sm-12">
                            <div class="title-line">
                                Client Information
                            </div>
                            <div class="table-row">
                                <div class="row">
                                    <div class="col-sm-3">Prefix</div>
                                    <div class="col-sm-9">
                                        
			<select size="1" name="Prefix">
                <option <% If Session("ClientPrefix") = "" Then Response.Write "selected" %>>Please select (optional)...</option>
                <option <% If Session("ClientPrefix") = "Mr." Then Response.Write "selected" %>>Mr.</option>
                <option <% If Session("ClientPrefix") = "Ms." Then Response.Write "selected" %>>Ms.</option>
                <option <% If Session("ClientPrefix") = "Dr." Then Response.Write "selected" %>>Dr.</option>
                <option <% If Session("ClientPrefix") = "Prof." Then Response.Write "selected" %>>Prof.</option>
            </select>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">First Name *</div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="FirstName" size="20" value="<% Response.Write Session("ClientFirstName") %>" maxlength="20" <% If Session("ClientAddress") = "Yes" And Session("ClientFirstName") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Last Name *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="LastName" size="20" value="<% Response.Write Session("ClientLastName") %>" maxlength="20" <% If Session("ClientAddress") = "Yes" And Session("ClientLastName") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Company/Organization *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="Company" size="40" value="<% Response.Write Session("ClientCompany") %>" maxlength="50">
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Street address 1 *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="StreetAddress" size="40" value="<% Response.Write Session("ClientStreetAddress") %>" maxlength="40"
          <% If Session("ClientAddress") = "Yes" And Session("ClientStreetAddress") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Street address 2 (optional)
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="StreetAddress2" size="40" value="<% Response.Write Session("ClientStreetAddress2") %>" maxlength="40">
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        City *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="City" size="20" value="<% Response.Write Session("ClientCity") %>" maxlength="20" <% If Session("ClientAddress") = "Yes" And Session("ClientCity") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <%
If Not Session("ShowClient") Then
                                %>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Country *
                                    </div>
                                    <div class="col-sm-9">
                                        
        <select name="Country" id="Country" onchange='setStates();' <% If Session("ClientAddress") = "Yes" And Session("ClientCountry") = "Please select a country..." Then Response.Write "class='txtBackground'" %>>
            <option>Please select a country...</option>
            <option <% If Session("ClientCountry") = "USA" Then Response.Write "selected" %>>USA</option>
            <option <% If Session("ClientCountry") = "Canada" Then Response.Write "selected" %>>Canada</option>

            <%
        Set objRec = Server.CreateObject ("ADODB.Recordset")
    	objRec.Open "Select Country From ShippingCalculator ORDER BY Country ASC", shippingcalculatorPath

        If Not objRec.EOF then
            Do While Not objRec.EOF
		If objRec("Country") <> "Canada" And objRec("Country") <> "USA" Then
		    If Session("ClientCountry") = objRec("Country") Then
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
                                        
	  <input type="text" name="Country" style="background-color: #FFF; border: none;" size="20" value="<%= Session("ClientCountry") %>"
          readonly>
                                    </div>
                                </div>
                                <% End If %>

                                <% If Session("ClientAddress") <> "Yes" And Session("ClientCountry") <> "USA" And Session("ClientCountry") <> "Canada" Then %>

                                <div class="row">
                                    <div class="col-sm-3">
                                        Province/State *
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="Province" id="Province" size="20" value="<%=Session("ClientProvince")%>" <% If Session("ClientAddress") = "Yes" And Session("ClientProvince") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>

                                <%
Else
DisplayProvinceDropdown Session("ClientCountry"), Session("ClientProvince")

End If
                                %>
                                <div class="row">
                                    <div class="col-sm-3">

                                        <%
                Select Case Session("ClientCountry")
                    Case "USA"
                        Response.Write "Zip code"
		    Case Else
                        Response.Write "Postal code"
                End select
                                        %>
             *
                                        
                                    </div>
                                    <div class="col-sm-9">
                                        
	    <input type="text" name="PostalCode" size="20" value="<% Response.Write Session("ClientPostalCode") %>" maxlength="15" <% If Session("ClientAddress") = "Yes" And Session("ClientPostalCode") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Email *
                                    </div>
                                    <div class="col-sm-9">
                                        
	    <input type="text" name="Email" size="40" value="<% Response.Write Session("ClientEmail") %>" maxlength="65" <% If Session("ClientAddress") = "Yes" And Session("ClientEmail") = "" Then Response.Write "class='txtBackground'" %>>
                                        <% If isTrial Then Response.Write "<div class ='email-restriction'>At this time, we cannot accept Hotmail, Gmail, or other free email providers for the trial license. If you absolutely need to use this address, please call us at (519)767-0197 or <a href='mailto:info@chiwater.com'>email us</a>.</div>" %>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Telephone *<br>
                                        (e.g. 123-123-1234)
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="Telephone" size="20" value="<% Response.Write Session("ClientTelephone") %>" maxlength="25" <% If Session("ClientAddress") = "Yes" And Session("ClientTelephone") = "" Then Response.Write "class='txtBackground'" %>>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Company website<br>
                                        (e.g. www. ... .com)
                                    </div>
                                    <div class="col-sm-9">
                                        
	  <input type="text" name="WebSite" size="40" value="<% Response.Write Session("ClientWebSite") %>" maxlength="65">
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Type of organization (choose closest)
                                    </div>
                                    <div class="col-sm-9">
                                        
                    <select size="1" name="Type">
                        <option <% If Session("ClientType") = "" Then Response.Write "selected" %>>Please select...</option>
                        <option <% If Session("ClientType") = "CE" Then Response.Write "selected" %>>Consulting Company</option>
                        <option <% If Session("ClientType") = "MA" Then Response.Write "selected" %>>Municipality</option>
                        <option <% If Session("ClientType") = "GA" Then Response.Write "selected" %>>Government Agency</option>
                        <option <% If Session("ClientType") = "OR" Then Response.Write "selected" %>>Organization/Association</option>
                        <option <% If Session("ClientType") = "UN" Then Response.Write "selected" %>>University/College</option>
                        <option <% If Session("ClientType") = "ST" Then Response.Write "selected" %>>Student</option>
                        <option <% If Session("ClientType") = "IN" Then Response.Write "selected" %>>Commercial/Industrial</option>
                    </select>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        <% = Question %>
                                    </div>
                                    <div class="col-sm-9">
                                        
		<select name="Source">
            <option>Please select...</option>
            <option>Email from CHI</option>
            <option>Colleague</option>
            <option>Web search</option>
            <option>Social Media</option>
            <option>Anthropocene Engineering Sales Representative</option>
        </select>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Stay connected with CHI?
                                    </div>
                                    <div class="col-sm-9">
                                        
            <select name="Consent">
                <option>Please select...</option>
                <option <% If Session("Consent") = 1 Then Response.Write "selected" %>>Yes, please keep me informed</option>
                <option <% If Session("Consent") = 0 Then Response.Write "selected" %>>No thanks</option>

            </select>

                                    </div>
                                </div>
                            </div>
                            <% If isConference Then %>
                            <div class="title-line">
                                Please share what conference topics are of most interest to you
                            </div>
                            <div class="table-row">
                                <div class="row">
                                    <div class="col-sm-3">
                                        Interest 1:
                                    </div>
                                    <div class="col-sm-9">
                                        
                <select name="Interest1">
                    <option>Please select...</option>
                    <option>2D modeling</option>
                    <option>Agriculture and BMPs</option>
                    <option>Calibration and optimization</option>
                    <option>Climate change</option>
                    <option>CSOs/SSOs</option>
                    <option>Field monitoring</option>
                    <option>Forecasting and decision analysis systems</option>
                    <option>Green infrastructure/LID/SUD</option>
                    <option>Habitat restoration</option>
                    <option>Integrated groundwater</option>
                    <option>Mining</option>
                    <option>New system design</option>
                    <option>Numerical methods</option>
                    <option>Precipitation / remote sensing</option>
                    <option>Real-time control</option>
                    <option>Waste water</option>
                    <option>Water quality</option>
                    <option>Water supply</option>
                </select>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Interest 2:
                                    </div>
                                    <div class="col-sm-9">
                                        
                <select name="Interest2">
                    <option>Please select...</option>
                    <option>2D modeling</option>
                    <option>Agriculture and BMPs</option>
                    <option>Calibration and optimization</option>
                    <option>Climate change</option>
                    <option>CSOs/SSOs</option>
                    <option>Field monitoring</option>
                    <option>Forecasting and decision analysis systems</option>
                    <option>Green infrastructure/LID/SUD</option>
                    <option>Habitat restoration</option>
                    <option>Integrated groundwater</option>
                    <option>Mining</option>
                    <option>New system design</option>
                    <option>Numerical methods</option>
                    <option>Precipitation / remote sensing</option>
                    <option>Real-time control</option>
                    <option>Waste water</option>
                    <option>Water quality</option>
                    <option>Water supply</option>
                </select>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Interest 3:
                                    </div>
                                    <div class="col-sm-9">
                                        
                <select name="Interest3">
                    <option>Please select...</option>
                    <option>2D modeling</option>
                    <option>Agriculture and BMPs</option>
                    <option>Calibration and optimization</option>
                    <option>Climate change</option>
                    <option>CSOs/SSOs</option>
                    <option>Field monitoring</option>
                    <option>Forecasting and decision analysis systems</option>
                    <option>Green infrastructure/LID/SUD</option>
                    <option>Habitat restoration</option>
                    <option>Integrated groundwater</option>
                    <option>Mining</option>
                    <option>New system design</option>
                    <option>Numerical methods</option>
                    <option>Precipitation / remote sensing</option>
                    <option>Real-time control</option>
                    <option>Waste water</option>
                    <option>Water quality</option>
                    <option>Water supply</option>
                </select>
                                    </div>
                                </div>
                            </div>
                            <% End If 

        ' workshop interests
        If isTraining And Not isWebTraining Then
                            %>
                            <div class="title-line brown">
                                Please share your modeling experience and what you would like to learn about in this workshop
                            </div>
                            <div class="table-row">
                                <div class="row">
                                    <div class="col-md-10">
                                        How much experience do you have with hydrology/hydraulic modeling software (not necessarily SWMM based)?
                                    </div>
                                    <div class="col-md-2">
                                        <select name="WExperience1">
                                            <option>Please select...</option>
                                            <option value="1">None</option>
                                            <option value="2">0 - 1 year</option>
                                            <option value="3">1 - 3 years</option>
                                            <option value="4">3 - 5 years</option>
                                            <option value="5">5 + years</option>
                                        </select>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-md-10">
                                        How much experience do you have with EPA SWMM5 or PCSWMM?
                                    </div>
                                    <div class="col-md-2">
                                        <select name="WExperience2">
                                            <option>Please select...</option>
                                            <option value="1">None</option>
                                            <option value="2">0 - 1 year</option>
                                            <option value="3">1 - 3 years</option>
                                            <option value="4">3 - 5 years</option>
                                            <option value="5">5 + years</option>
                                        </select>
                                    </div>
                                </div>

                            </div>

<% If (Session("AskWorkshopInterests") Or Not isSpecialty) Then %>
                            <div class="title-line brown">
                                Please share what workshop topics are of interest to you:
                            </div>
                            <div class="table-row">
                                <div class="row">
                                    <div class="col-sm-3">
                                        Interest 1:
                                    </div>
                                    <div class="col-sm-9">
                     
                <select name="WInterest1">
                    <option>Please select...</option>
                    <% If (DayAll Or DayOne) Then %>
                    <option value="1">Introduction to PCSWMM/SWMM5</option>
                    <option value="17">SWMM5 Hydrology</option>
                    <option value="18">SWMM5 Hydraulics</option>
                    <option value="19">Automated Model Parameterization in PCSWMM</option>
                    <option value="15">Python Scripting in PCSWMM</option>
                    <option value="5">SWMM5 Wastewater Collection Systems</option>
                    <% End If
                       If (DayAll Or DayTwo) Then %>
                    <option value="8">SWMM5 Water Quality</option>
                    <option value="9">SWMM5 Low Impact Development</option>
                    <option value="3">Integrated 1D-2D Modeling in PCSWMM</option>
                    <option value="13">Watershed Modeling in PCSWMM</option>
                    <option value="20">Urban Flooding Analysis in PCSWMM</option>
                    <option value="12">Duration Analysis in PCSWMM</option>
                    <option value="10">Calibration and Error Analysis in PCSWMM</option>
                    <% End If %>

                    <!--<option value="2">Floodplain inundation </option>-->
                    <!--<option value="4">Dual drainage (major/minor system)</option>-->
                    <!--<option value="6">Rural and/or watershed applications</option>-->
<!--                    <option value="7">Importing from HEC-RAS</option>
                    <option value="11">Time series management and analysis</option>
                    <option value="14">Agricultural modeling</option>
                    <option value="16">Radar Acquisition and Processing</option>-->

                </select>

                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Interest 2:
                                    </div>
                                    <div class="col-sm-9">
                                        
                <select name="WInterest2">
                    <option>Please select...</option>
                    <% If (DayAll Or DayOne) Then %>
                    <option value="1">Introduction to PCSWMM/SWMM5</option>
                    <option value="17">SWMM5 Hydrology</option>
                    <option value="18">SWMM5 Hydraulics</option>
                    <option value="19">Automated Model Parameterization in PCSWMM</option>
                    <option value="15">Python Scripting in PCSWMM</option>
                    <option value="5">SWMM5 Wastewater Collection Systems</option>
                    <% End If
                       If (DayAll Or DayTwo) Then %>
                    <option value="8">SWMM5 Water Quality</option>
                    <option value="9">SWMM5 Low Impact Development</option>
                    <option value="3">Integrated 1D-2D Modeling in PCSWMM</option>
                    <option value="13">Watershed Modeling in PCSWMM</option>
                    <option value="20">Urban Flooding Analysis in PCSWMM</option>
                    <option value="12">Duration Analysis in PCSWMM</option>
                    <option value="10">Calibration and Error Analysis in PCSWMM</option>
                    <% End If %>
                </select>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="col-sm-3">
                                        Interest 3:
                                    </div>
                                    <div class="col-sm-9">
                                        
                <select name="WInterest3">
                    <option>Please select...</option>
                    <% If (DayAll Or DayOne) Then %>
                    <option value="1">Introduction to PCSWMM/SWMM5</option>
                    <option value="17">SWMM5 Hydrology</option>
                    <option value="18">SWMM5 Hydraulics</option>
                    <option value="19">Automated Model Parameterization in PCSWMM</option>
                    <option value="15">Python Scripting in PCSWMM</option>
                    <option value="5">SWMM5 Wastewater Collection Systems</option>
                    <% End If
                       If (DayAll Or DayTwo) Then %>
                    <option value="8">SWMM5 Water Quality</option>
                    <option value="9">SWMM5 Low Impact Development</option>
                    <option value="3">Integrated 1D-2D Modeling in PCSWMM</option>
                    <option value="13">Watershed Modeling in PCSWMM</option>
                    <option value="20">Urban Flooding Analysis in PCSWMM</option>
                    <option value="12">Duration Analysis in PCSWMM</option>
                    <option value="10">Calibration and Error Analysis in PCSWMM</option>
                    <% End If %>
                </select>
                                    </div>
                                </div>
                            </div>
<% End If %>

                            <% End If %>
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
                                    <% If Session("FormCheckOut") <> "Yes" Then 
            If Not  isTrial Then
                                    %>
                                    <button type="submit" form="ClientAddressForm" class="btn btn-blue">
                                        NEXT
                                    <img class="arrow" src="assets/arrow-right.png" /></button>
                                    <% Else %>
                                    <button type="submit" form="ClientAddressForm" class="btn btn-blue">
                                        SUBMIT ORDER
                                    <img class="arrow" src="assets/arrow-right.png" /></button>
                                    <%      End If
        Else %>
                                    <button type="submit" form="ClientAddressForm" class="btn btn-blue">
                                        UPDATE
                                    <img class="arrow" src="assets/arrow-right.png" /></button>
                                    <a href="checkout.asp" class="btn btn-blue">CANCEL
                                    <img class="arrow" src="assets/arrow-right.png" /></a>
                                    <% End If %>
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
