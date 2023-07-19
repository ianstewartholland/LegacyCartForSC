<!--#include virtual ="/chi/includes/shippingcalculator.asp"-->
<!--#include virtual ="/chi/includes/taxcalculator.asp"-->
<!--#include file="LogFile.asp"-->

<%
	'check to ensure CartItems exist
	If Session("NumCartItems") > 0 Then
		'everything ok
	Else
		Response.Redirect("cart.asp")
	End If
        If Request.QueryString("quote") = "yes" Then
            Session("ClientQuote") = True
        Else
            Session("ClientQuote") = False
        End If

	'check which next page to redirect
	If Session("PayQuoteMode") = "Yes" Or Trim(Request.QueryString("Direct")) = "Yes" Then
	    ' okay
	Else
	    If Not Session("AddressClient") Then Response.Redirect "clientaddress.asp"
	End If

	Session("FormCheckOut") = "Yes"
	Session.Timeout = 60

	'
	' if there is shipping item, use ship address to calculate; if not, use client address to calculate
	'
	' if has not change shipping address and pay quote, do not recalculate
	ShowFood = False
	ShowDay2 = False
    ChangePricing = False
    If IsNull(Session("ChangePricing")) Then Session("ChangePricing") = False

	If Session("ShipCountry") <> "" And Session("ShipProvince") <> "" And Session("SubTotalCost") > 0 Then
	    If Session("PayQuoteMode") <> "Yes" Or Session("RecalculateQuote") = "Yes" Then
		Session("ShippingCost") = 0
		Session("ShippingRequired") = ""

		hasShipping = False
		isTShirt = False
                isWebTraining = False
		WebTrainingItem = "web workshop"
		ItemStr = ""
		For t = 1 to Session("NumCartItems")
		    ItemStr = ItemStr & Session("CartItem(" & t & ")") & "/" & Session("CartItemDescription(" & t & ")") & "/" & Session("CartItemCost(" & t & ")") & "/" & Session("CartItemQuantity(" & t & ")") & "/"
		    ' check web workshop location, item number contains book R
		    ' no longer check international web workshops because our code does not handle the international shipping cost
		    ' todo - should have mentioned this on the webpage
		    ' Or (InStr(Session("CartItemDescription(" & t & ")"), WebTrainingItem) > 0 And (Session("ShipCountry") <> "Canada" Or Session("ShipCountry") <> "USA"))
		    If (LCase(Left(Session("CartItem(" & t & ")"), 1)) = "r" And Len(Session("CartItem(" & t & ")")) < 5) Or (InStr(Session("CartItemDescription(" & t & ")"), WebTrainingItem) > 0 And Session("ShipCountry") <> "Canada" And Session("ShipCountry") <> "USA") Then hasShipping = True

                    If InStr(Session("CartItemDescription(" & t & ")"), WebTrainingItem) > 0 Then isWebTraining = True

		    ' when I setup the page, indian and france workshops are not in options but change if it changes
		    ' show special client notes for conference or north america workshops
		    If LCase(Left(Session("CartItem(" & t & ")"), 1)) = "c" Or LCase(Left(Session("CartItem(" & t & ")"), 1)) = "w" Then
			ShowFood = True
		    End If
		    ' And Len(Session("CartItem(" & t & ")")) = 4
		    If LCase(Left(Session("CartItem(" & t & ")"), 1)) = "w" And InStr(Session("CartItemDescription(" & t & ")"), WebTrainingItem) = 0 And InStr(Session("CartItemDescription(" & t & ")"), "South Africa") = 0 Then
			ShowDay2 = True
		    End If

		    ' t-shirt
		    If Session("CartItem(" & t & ")") = "A001" Then isTShirt = True

            ' check if conference and toronto workshop attendee
            If Left(Session("CartItem(" & t & ")"), 4) = "C031" Or InStr(Session("CartItem(" & t & ")"), "W1156") > 0 Then   
                'ChangePricing = True
            End If

		Next

    If isWebTraining And hasShipping Then hasShipping = False

		If hasShipping Or isTShirt Then

		    If Session("PayQuoteMode") <> "Yes" Then
				' prevent change shipping cost for quotes or orders
				If Not isTShirt Then
					Session("ShippingCost") = CalculateShipping (ItemStr, Session("ShipCountry"), Session("ShipProvince"), Session("ShipCity"), Session("ShipPostalCode"), Session("Currency"))
				Else
					Session("ShippingCost") = 18
				End If
				Session("ShippingRequired") = "Yes"
		    End If
		End If

		' ini
		If IsNull(Session("NoGST")) Then Session("NoGST") = ""
		If IsNull(Session("NoPST")) Then Session("NoGST") = ""
		If IsNull(Session("NoHST")) Then Session("NoGST") = ""

		Tax = TaxCalculator (Session("ShipProvince"), Session("CurrencyAbrev"), ItemStr, Session("ShippingCost"), Session("NoGST"), Session("NoPST"), Session("NoHST"))

		Session("GSTCost") = Tax(0)
		Session("PSTCost") = Tax(1)
		Session("HSTCost") = Tax(2)
		Session("TotalCost") = Tax(3)
		Session("SubTotalCost") = Tax(4)
	    End If
	End If

	If Session("SubTotalCost") = 0 Then
	    Session("TotalCost") = 0
	End If

  	Session("LastPage") = "checkout.asp"

	' check payment form
	' we will try to be generous to our clients so we do not care their address anymore as long as they give us shipping country and province, we let them proceed
	proceed = False
	If (Session("SubTotalCost") > 0 And Session("PaymentForm") = "Yes" And Session("ShipCountry") <> "") Or (Session("SubTotalCost") = 0 And Session("ShowClient")) Or (Session("SubTotalCost") > 0 And Session("PayQuoteMode") = "Yes" ) Or (Session("SubTotalCost") > 0 And Session("CardType") = "Wire transfer" And Session("PayQuoteMode") = "Yes" ) Then proceed = True

    ' update US pricing for 2016 and onwards conference and toronto workshop 
'    If Session("CurrencyAbrev") = "US" And ChangePricing And Not Session("PayQuoteMode") = "Yes" And Not Session("ChangePricing") Then
'        Session("SubTotalCost") = 0
'        For t = 1 to Session("NumCartItems")
'            'If Session("CartItem(" & t & ")") <> "W955X" Then 
'                Session("CartItemCost(" & t & ")") = Round(Session("CartItemCost(" & t & ")") / 1.25/5)*5
'            'End If
'            Session("SubTotalCost") = Session("SubTotalCost") + Session("CartItemCost(" & t & ")") 
'        Next
'        Session("SubTotalCost") = CCur(Session("SubTotalCost"))
'		Session("HSTCost") = CCur(Session("SubTotalCost") * 0.13)
'        Session("TotalCost") = Session("SubTotalCost") + Session("HSTCost")
'        Session("ChangePricing") = True
'    End If

	' check for conference early bird discount if pass the deadline - check every year to avoid mistakes - 2012/01/27
    ' check C252X - 2016
    showConferenceInterests = False
	For t = 1 to Session("NumCartItems")
	    If DateDiff("d", Date, CDate("2022/12/31")) < 0 Then
		    If Len(Session("CartItem(" & t & ")")) = 5 And Session("CartItem(" & t & ")") <> "C312X" And Left(Session("CartItem(" & t & ")"), 1) = "C" And Right(Session("CartItem(" & t & ")"), 1) = "X" And (Not IsNull(Session("CurrentTransaction")) And Left(Session("CurrentTransaction"), 1) = "Q")  Then
		        If InStr(Session("CartItemDescription(" & t & ")"), " - Expired") < 1 Then
                    Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & " - Expired"
                End If
		        Session("SubTotalCost") = Session("SubTotalCost") - Session("CartItemCost(" & t & ")")
		        Session("HSTCost") = Session("SubTotalCost") * 0.13
		        Session("TotalCost") = Session("SubTotalCost") + Session("HSTCost")
		        Session("CartItemCost(" & t & ")") = 0
		    End If
	    End If
            
            If Session("CartItem(" & t & ")") = "C031" And IsEmpty(Session("ConferenceInterests")) Then
                showConferenceInterests = True
            End If 
	Next

    If showConferenceInterests Then Response.Redirect("clientaddress.asp")

%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Order Summary - CHI Shopping Cart</title>
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
                        <h1>Order Summary</h1>
                        <p>
                            <% If proceed Then Response.Write "Your order is ready for submission. " %>
                            Please review the information on this page carefully. If you wish to make any
            changes to the client<% If Session("SubTotalCost") > 0 Then Response.Write ", shipping or payment" %> information, click on <i>
                Edit</i>. <% If Session("SubTotalCost") > 0 Then Response.Write "To change the items in the cart, click on the <i>Edit Cart</i> button. " %>When
                            you are ready, click on the <i>Submit Order</i> button to place your order.
            You will receive a confirmation by email when the order is processed by CHI staff.
                        </p>
                    </div>
                </div>
            </div>
        </div>

        <div class="container-fluid white">
            <div class="container checkout">
                <div class="row">
                    <div class="col-sm-12 ">
                        <div class="title-line">
                            Client Information &nbsp;|&nbsp; 
     <%
	    If Session("PayQuoteMode") = "Yes" Then
		Response.Write "Please <a href = 'mailto:info@chiwater.com?Subject=Change address for order (" & Session("CurrentTransaction") & ")&body=Please let us know your updated address. We will update our records as soon as possible. %0D%0A"
		Response.Write "'>email us</a> if you need to change your address for this order."
	    Else %>
                            <a href="clientaddress.asp">EDIT
                                <img class="arrow" src="assets/arrow-right-teal.png" /></a>
                            <% End If %>
                        </div>
                        <div>
                            <%
	If Session("PayQuoteMode") <> "Yes" And Not Session("ShowClient") Then
                            %>
                            <p class="error">Please enter client information before proceeding.</p>
                            <% Else %>
                            <div class="row">
                                <div class="col-sm-4">
                                    <%
            Response.Write "<p><span class='subtitle'>" & Session("ClientName") & "</span><br>"
            If Session("ClientCompany") <> "" Then Response.Write Session("ClientCompany") & "<br>"
            Response.Write Session("ClientStreetAddress") & "<br>"
            If Session("ClientStreetAddress2") <> "" Then Response.Write Session("ClientStreetAddress2") & "<br>"
            Response.Write Session("ClientCity") & " " & Session("ClientProvince") & " " & Session("ClientPostalCode") & "<br>"
            Response.Write Session("ClientCountry") & "</p>"
                                    %>
                                </div>
                                <div class="col-sm-8">
                                    <%
            Response.Write "<p>Email: " & Session("ClientEmail") & "<br>"
            Response.Write "Tel: " & Session("ClientTelephone") & "<br>"
            If Session("ClientCellPhone") <> "" Then Response.Write "Cell: " & Session("ClientCellPhone") & "<br>"
            If Session("ClientFax") <> "" Then Response.Write "Fax: " & Session("ClientFax") & "<br>"
            If Session("ClientWebSite") <> "" Then Response.Write "Web: " & Session("ClientWebSite") & "</p>"
                                    %>
                                </div>
                            </div>
                            <% End If %>
                        </div>
                    </div>
                </div>

                <!-- billing - no shipping -->
                <% If Session("SubTotalCost") > 0 Then %>
                <div class="row">
                    <div class="col-sm-12 ">
                        <div class="title-line">
                            Billing Address &nbsp;|&nbsp; 
     <% If Session("PayQuoteMode") <> "Yes" Then %>
                            <a href="billingaddress.asp">EDIT
                                <img class="arrow" src="assets/arrow-right-teal.png" /></a>
                            <% End If %>
                        </div>
                        <div>
                            <%
        'check if billing address is the same as the client address
        BillingSame = "Yes"
        If Session("BillName") <> Session("ClientName") Then BillingSame = "No"
        If Session("BillCompany") <> Session("ClientCompany") Then BillingSame = "No"
        If Session("BillStreetAddress") <> Session("ClientStreetAddress") Then BillingSame = "No"
        If Session("BillStreetAddress2") <> Session("ClientStreetAddress2") Then BillingSame = "No"
        If Session("BillCity") <> Session("ClientCity") Then BillingSame = "No"
        If Session("BillProvince") <> Session("ClientProvince") Then BillingSame = "No"
        If Session("BillPostalCode") <> Session("ClientPostalCode") Then BillingSame = "No"
        If Session("BillCountry") <> Session("ClientCountry") Then BillingSame = "No"
        If Session("BillEmail") <> Session("ClientEmail") Then BillingSame = "No"
        If Session("BillTelephone") <> Session("ClientTelephone") Then BillingSame = "No"
        If Session("BillFax") <> Session("ClientFax") Then BillingSame = "No"
        If BillingSame = "Yes" Then
                            %>
                            <p>Same as client address above...</p>
                            <%
        Else
                            %>
                            <div class="row">
                                <div class="col-sm-4">
                                    <%
                        Response.Write "<p><span class='subtitle'>" & Session("BillName") & "</span><br>"
                        If Session("BillCompany") <> "" Then Response.Write Session("BillCompany") & "<br>"
                        Response.Write Session("BillStreetAddress") & "<br>"
                        If Session("BillStreetAddress2") <> "" Then Response.Write Session("BillStreetAddress2") & "<br>"
                        Response.Write Session("BillCity") & " " & Session("BillProvince") & " " & Session("BillPostalCode") & "<br>"
                        Response.Write Session("BillCountry") & "</p>"
                                    %>
                                </div>
                                <div class="col-sm-8">
                                    <%
                        Response.Write "<p>Email: " & Session("BillEmail") & "<br>"
                        Response.Write "Tel: " & Session("BillTelephone") & "<br>"
                        If Session("BillFax") <> "" Then Response.Write "Fax: " & Session("BillFax") & "</p>"
                                    %>
                                </div>
                            </div>
                            <%
        End If
                            %>
                        </div>
                    </div>
                </div>
                <% 'End If %>

                <!-- payment method -->
                <% If (Session("PayQuoteMode") <> "Yes" And Session("PaymentForm") = "Yes") Or (Session("PayQuoteMode") = "Yes" And Session("CardType") <> "") Then %>
                <div class="row">
                    <div class="col-sm-12 ">
                        <div class="title-line">
                            Payment Method &nbsp;|&nbsp;
	<%
	' prevent people change an already posted order
	If Session("PayQuoteMode") = "Yes" And Session("IsPosted") = "Yes" And Session("CardType") = "Wire transfer" Then
	    Response.Write "Please <a href = 'mailto:info@chiwater.com?Subject=Change payment method for order (" & Session("CurrentTransaction") & ")&body=Please let us know your new payment method. We will handle as soon as possible. %0D%0A"
	    Response.Write "'>email us</a> if you need to change your payment method for this order."
	Else %>
                            <a href="paymentform.asp">EDIT
                                <img class="arrow" src="assets/arrow-right-teal.png" /></a>
                            <% End If %>
                        </div>
                        <div>
                            <p>
                                <%
            If Session("CardType") = "Invoice" Then
                                %>Please invoice...<%
            ElseIf Session("CardType") = "Wire transfer" Then
                                %>Wire transfer in <b>US dollars</b> to:
                            </p>
                            <p style='padding-left: 0px;'>
                                <b>Beneficiary bank:</b>
                                <br>
                                CANADIAN IMPERIAL BANK OF COMMERCE<br>
                                59 WYNDHAM ST N<br>
                                GUELPH, ONTARIO<br>
                                CANADA<br>
                                <br>
                                Swift # CIBCCATT
                                <br>
                                Transit # 00052
                                <br>
                                Institution # 0010
                                <br>
                                Account #: 0266310
                                <br>
                                Account name: Computational Hydraulics Inc.
                            </p>
                            <p>
                                A fee of $20 has been added to cover bank charges. 
                Please <i>Submit Order</i> and then <b>email a PDF proof of payment</b> to
                <a href="mailto:info@chiwater.com?subject=Wire transfer">info@chiwater.com</a><%
            Else
                        Response.Write "Card type: " & Session("CardType") & "<br>"
                        Response.Write "Card number: " & HideCreditCardNumber(MakeNumeric(Session("CardNumber"))) & "<br>"
                        Response.Write "Card expiry: " & Session("CardExpiry") & "<br>"
                        Response.Write "Cardholder name: " & Session("CardName")
            End If
                %>
                            </p>
                        </div>
                    </div>
                </div>
                <% End If %>

                <!-- purchase order -->
                <div class="row">
                    <div class="col-sm-12 ">
                        <div class="title-line">
                            Purchase Order No. &nbsp;|&nbsp;
	<% If Len(Session("PurchaseOrder")) > 1 then  %>
                            <a href="purchaseorder.asp">EDIT<img class="arrow" src="assets/arrow-right-teal.png" /></a>
                            <% Else %>
                            <a href="purchaseorder.asp">Click here to add a purchase order number to this order<img class="arrow" src="assets/arrow-right-teal.png" /></a>
                            <% End If %>
                        </div>
                        <div>
                            <p>
                                <%
            If Len(Session("PurchaseOrder")) > 1 then
                Response.Write Session("PurchaseOrder") 
            End If
                                %>
                            </p>
                        </div>
                    </div>
                </div>

                <% End If %>

                <!-- client notes -->
                <div class="row">
                    <div class="col-sm-12 ">
                        <div class="title-line">
                            Client Notes &nbsp;|&nbsp;
	<% If Len(Session("ClientNotes")) > 1 then  %>
                            <a href="clientnotes.asp">EDIT<img class="arrow" src="assets/arrow-right-teal.png" /></a>
                            <% Else %>
                            <a href="clientnotes.asp">Click here to add a comment to this order.
			<% If isTShirt Then Response.Write "For example, please specify the T-shirt size. " %>
                                <% If ShowFood Then Response.Write "For example, please let us know if you have any dietary requirements. " %>
                                <% If ShowDay2 Then Response.Write "Also please indicate which day 2 subjects are of most interest (e.g. LID, 2D, sanitary, calibration or HEC-RAS importing)." %>
                            </a>
                            <% End If %>
                        </div>
                        <div>
                            <p>
                                <%
            If Len(Session("ClientNotes")) > 1 then
                Response.Write Session("ClientNotes") 
            End If
                                %>
                            </p>
                        </div>
                    </div>
                </div>

                <!-- cart -->
                <div class="row">
                    <div class="col-sm-12">
                        <div class="title-line">
                            Shopping Cart &nbsp;|&nbsp;
                            <a href="cart.asp">EDIT<img class="arrow" src="assets/arrow-right-teal.png" /></a>
                        </div>

                        <div class="table-row shopping-cart">
                            <div class="row title-line adjust-spacing">
                                <div class="col-sm-1 col-xs-2">
                                    ITEM
                                </div>
                                <div class="col-md-7 col-sm-5 col-xs-10">
                                    DESCRIPTION
                                </div>
                                <div class="col-sm-2 right hidden-xs">
                                    UNIT PRICE
                                </div>
                                <div class="col-md-1 col-sm-2 hidden-xs right">
                                    QUANTITY
                                </div>
                                <div class="col-md-1 col-sm-2 hidden-xs">
                                    AMOUNT
                                </div>
                            </div>
                            <%
          For t = 1 to Session("NumCartItems")
                  TotalItemCost = (Session("CartItemCost(" & t & ")") * Session("CartItemQuantity(" & t & ")"))
                            %>
                            <div class="row">
                                <div class="col-sm-1 col-xs-2">
                                    <% response.write Session("CartItem(" & t & ")") %>
                                </div>
                                <div class="col-md-7 col-sm-5 col-xs-10">
                                    <% response.write Session("CartItemDescription(" & t & ")") %>
                                </div>
                                <div class="col-sm-2 right hidden-xs">
                                    <% response.write FormatCurrency(Session("CartItemCost(" & t & ")"),2) %>
                                </div>
                                <div class="col-md-1 col-sm-2 col-xs-12 right">
                                    <div class="visible-xs-inline"><% response.write FormatCurrency(Session("CartItemCost(" & t & ")"),2) %> <b class="purple">*</b>
                                    </div>
                                    <% response.write Session("CartItemQuantity(" & t & ")") %>
                                </div>
                                <div class="col-md-1 col-sm-2 hidden-xs">
                                    <% response.write FormatCurrency(TotalItemCost,2) %>
                                </div>
                            </div>
                            <% Next %>
                        </div>

                        <div class="row summary-line">
                            <div class="row">
                                <div class="col-md-11 col-sm-10 hidden-xs">
                                    Subtotal:
                                </div>
                                <div class="col-md-1 col-sm-2 col-xs-12">
                                    <div class="visible-xs-inline">Subtotal:</div>

                                    <% response.write FormatCurrency(Session("SubTotalCost"),2) %>
                                </div>
                            </div>
                            <% If Session("HSTCost") > 0 Then %>
                            <div class="row">
                                <div class="col-md-11 col-sm-10 hidden-xs">HST (#119880979):</div>
                                <div class="col-md-1 col-sm-2 col-xs-12">
                                    <div class="visible-xs-inline">HST (#119880979):</div>
                                    <% response.write FormatCurrency(Session("HSTCost"),2) %>
                                </div>
                            </div>
                            <%
                End If
  		If Session("GSTCost") > 0 Then
                            %>
                            <div class="row">
                                <div class="col-md-11 col-sm-10 hidden-xs">GST (#119880979):</div>
                                <div class="col-md-1 col-sm-2 col-xs-12">
                                    <div class="visible-xs-inline">GST (#119880979):</div>
                                    <% response.write FormatCurrency(Session("GSTCost"),2) %>
                                </div>
                            </div>
                            <%
  		End If
                            %>
                            <div class="row purple">
                                <div class="col-md-11 col-sm-10 hidden-xs"><b>Total (<% Response.Write Session("CurrencyAbrev") %>$):</b></div>
                                <div class="col-md-1 col-sm-2 col-xs-12">
                                    <div class="visible-xs-inline"><b>Total:</b></div>
                                    <b><% response.write FormatCurrency(Session("TotalCost"),2) %></b>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="container-fluid">
            <div class="container">
                <div class="row">
                    <div class="col-sm-12">
                        <div class="row button-row">
                            <div class="col-xs-12 right">
                                <% If Session("PayQuoteMode") <> "Yes" Then %>
                                <a href='quote.asp' class="btn btn-teal">PRINT QUOTE
                                    <img class="arrow" src="assets/arrow-right.png" /></a>
                                <% End If %>
                                <%
	     If (Session("PayQuoteMode") <> "Yes" And Session("PaymentForm") <> "Yes" And Session("SubTotalCost") > 0) Or (Session("PayQuoteMode") = "Yes" And (Session("CardType") = "" And Session("SubTotalCost") > 0))Then
		    Response.Write "<a href='paymentform.asp' class='btn btn-teal'>CHOOSE PAYMENT METHOD<img class='arrow' src='assets/arrow-right.png' /></a>"
	     Else
            Response.Write "<a class='btn btn-blue' href='"
		    If (Session("PayQuoteMode") <> "Yes" And Session("ShowClient")) Or (Session("PayQuoteMode") = "Yes" And Session("CardType") <> "") Or (Session("PayQuoteMode") = "Yes" And Session("SubTotalCost") = 0) Then Response.Write "ordercommit.asp" Else Response.Write "checkout.asp"
		    Response.Write "'>SUBMIT ORDER<img class='arrow' src='assets/arrow-right.png' /></a>"
	     End If
                                %>
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


    </div>

    <!--#include file ="assets/footer.asp"-->
</body>
</html>
