<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<%

	'check to ensure CartItems exist
	If Session("NumCartItems") > 0 Then
	    'everything ok
	    If Session("SubTotalCost") = 0 Then Response.Redirect("checkout.asp")
	    If Session("PayQuoteMode") = "Yes" And Session("IsPosted") = "Yes" And Session("CardType") = "Wire transfer" Then Response.Redirect("checkout.asp")
	Else
	    Response.Redirect("cart.asp")
	End If
	
	Session("FormPayment") = True
	
	'Session("ShowPayment") = True

	If Request.QueryString("Action") = "Do" Then
	    Session("PaymentForm") = ""
	    dim CardNumber
    
	    'check which drop down they used (method of payment)
	'    If Request.Form("CardType2") = "Wire transfer" Then
	'	Session("CardType") = "Wire transfer"
	'    ElseIf Request.Form("CardType2") = "Invoice" Then
	'	Session("CardType") = "Invoice"
	'    Else
	'	Session("CardType") = Request.Form("CardType")
	'    End If
	

	    '
	    ' rewrite 
	    '
	    Session("CardType2") = Request.Form("CardType2")
	    Select Case Request.Form("CardType2")
		Case "Wire transfer"
		    Session("CardType") = "Wire transfer"
		Case "Invoice"
		    Session("CardType") = "Invoice"
		Case Else
		    Session("CardType") = Request.Form("CardType")
		    If Left(Session("CardType"), 6) <> "Choose" Then
			Session("CardExpiry") = Request.Form("CardExpiry")
			Session("CardName") = Request.Form("CardName")
			
			CardNumber = Trim(CStr(Request.Form("CardNumber")))
			If Len(CardNumber) > 13 Then
			    If Instr(1,CardNumber,"-") = 0 And Instr(1,CardNumber," ") = 0 Then
				CardNumber = Mid(CardNumber,1,4) & " " & Mid(CardNumber,5,4) & " " & Mid(CardNumber,9,4) & " " & Mid(CardNumber,13)
			    End If
			End If
			Session("CardNumber") = CardNumber
		    Else
			Session("CardType") = ""
		    End If
	    End Select
	    

	    
	    'Session("CardExpiry") = Request.Form("CardExpiry")
	    'Session("CardName") = Request.Form("CardName")
	    
	    'if Left(Session("CardType"), 6) = "Choose" Then Session("CardType") = ""
    
	    'format card number
	    'CardNumber = Trim(CStr(Request.Form("CardNumber")))
    
	'    If Len(CardNumber) > 13 Then
	'	    If Instr(1,CardNumber,"-") = 0 And Instr(1,CardNumber," ") = 0 Then
	'		    CardNumber = Mid(CardNumber,1,4) & " " & Mid(CardNumber,5,4) & " " & Mid(CardNumber,9,4) & " " & Mid(CardNumber,13)
	'	    End If
	'    End If
	    'Session("CardNumber") = CardNumber
	    
	    
	    
	    'check to add or remove wire payment fee
	    'WireItem = "M301"
	    If Session("CardType") = "Wire transfer" Then
		    For t = 1 to Session("NumCartItems")
			    If Session("CartItem(" & t & ")") = WireItem Then Exit For
		    Next
		    If t > Session("NumCartItems") Then
			    'add wire fee
			'    For n = 1 to Session("NumCatalogItems")
			'	    'get item and number
			'	    If Session("CatalogItem(" & n & ")") = WireItem Then
			'		    t = Session("NumCartItems") + 1
			'		    Session("CartItem(" & t & ")") = WireItem
			'		    Session("CartItemQuantity(" & t & ")") = 1
			'		    Session("CartItemDescription(" & t & ")") = Session("CatalogItemDescription(" & n & ")")
			'		    Session("CartItemCost(" & t & ")") = Session("CatalogItemCost(" & n & ")")
			'		    Session("NumCartItems") = t
			'	    End If
			'	    'Add to the count of the total number of items
			'	    Session("TotalNumItems") = Session("TotalNumItems") + NumItems
			'    Next
			
			'
			' rewrite the above
			'
			' find out the wire item in the database
			'Set objRecCat = Server.CreateObject ("ADODB.Recordset")	
			'objRecCat.Open "Catalog", catalogPath, 3, 3, 2
			'objRecCat.Filter = "ItemNo = '" & WireItem & "'"
			'objRecCat.MoveFirst
			
			t = Session("NumCartItems") + 1
			Session("CartItem(" & t & ")") = WireItem
			Session("CartItemQuantity(" & t & ")") = 1
			Session("CartItemDescription(" & t & ")") = "Wire transfer fee (bank charge)" ' objRecCat("Description")
			
			'If Session("Currency") = "Canadian" Then
			'    Session("CartItemCost(" & t & ")") = objRecCat("CADCost")
			'Else
			    Session("CartItemCost(" & t & ")") = 20 'objRecCat("USCost")
			'End If

			Session("NumCartItems") = t
			
			If Session("PayQuoteMode") = "Yes" Then Session("RecalculateQuote") = "Yes"
			    'set flags for recalculating
			    'Session("SubTotalCalculated") = ""
			    'Session("ShippingCalculated") = ""
			    'Session("TaxCalculated") = ""
			    'Session("TotalCost") = 0
		    End If
		    
		    
		    
	    Else
		    For t = 1 to Session("NumCartItems")
			    If Session("CartItem(" & t & ")") = WireItem Then
				If Session("PayQuoteMode") = "Yes" Then Session("RecalculateQuote") = "Yes"
				Exit For
			    End If
		    Next
		    If t <= Session("NumCartItems") Then
			    'remove wire fee
			    For n = t to Session("NumCartItems") - 1
				    t = Session("NumCartItems") + 1
				    Session("CartItem(" & n & ")") = Session("CartItem(" & n + 1 & ")")
				    Session("CartItemQuantity(" & n & ")") = Session("CartItemQuantity(" & n + 1 & ")")
				    Session("CatalogItemDescription(" & n & ")") = Session("CatalogItemDescription(" & n + 1 & ")")
				    Session("CatalogItemCost(" & n & ")") = Session("CatalogItemCost(" & n + 1 & ")")
			    Next
			    Session("NumCartItems") = Session("NumCartItems") - 1
			    'Session("TotalNumItems") = Session("TotalNumItems") - 1
			    'set flags for recalculating
			    'Session("SubTotalCalculated") = ""
			    'Session("ShippingCalculated") = ""
			    'Session("TaxCalculated") = ""
			    'Session("TotalCost") = 0

		    End If
	    End If

	    

	    ' add validation
	    If Request.Form("CardType2") = "Choose payment type..." Then Response.Redirect "paymentform.asp?Error=Type"   
	    If Request.Form("CardType2") = "cc" And Request.Form("CardType") = "Choose card type..." Then Response.Redirect "paymentform.asp?Error=Card"
	    
	    If Request.Form("CardType2") = "cc" Then
		If Request.Form("CardName") = "" Or Request.Form("CardExpiry") = "" Or Request.Form("CardNumber") = "" Then Response.Redirect "paymentform.asp?Error=Fill"
	    End If
	    Session("PaymentForm") = "Yes"

	    Response.Redirect "checkout.asp"
	
	End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Payment Options - CHI Shopping Cart</title>
    <link rel="shortcut icon" href="assets/chilogo.ico" />
    <link href='https://fonts.googleapis.com/css?family=PT+Serif|Fjalla+One|Open+Sans:300italic,400italic,600italic,300,400,600,700'
        rel='stylesheet' type='text/css'>
    <link href="assets/site.min.css" rel="stylesheet" />
    <link href="assets/style.min.css" rel="stylesheet" />
    <script src="assets/startup.service.js"></script>
    <script src="assets/cart.js"></script>
    <script type="text/javascript">

        //on load
        function init() {
            if (document.form1.CardType2.value != "cc") {
                disableField();
            }
        }

        function disableField() {
            document.form1.CardType.disabled = true;
            document.form1.CardNumber.disabled = true;
            document.form1.CardExpiry.disabled = true;
            document.form1.CardName.disabled = true;
        }

        function enableField() {
            document.form1.CardType.disabled = false;
            document.form1.CardNumber.disabled = false;
            document.form1.CardExpiry.disabled = false;
            document.form1.CardName.disabled = false;
        }

        function check_text() {
            if (document.form1.CardType2.value == "cc") {
                enableField();
            }
            else {
                disableField();
            }
        }

        // validate form for wire transfer selection
        function validateForm() {
            var sel = document.getElementById("CardType2");

            if (sel != null) {
                if (sel.options[sel.selectedIndex].text == "Wire transfer")
                    return confirm("There is an additional $20 USD wire transfer fee that will be added to your order.\r\n\r\nIf you would like to proceed please choose OK.  If you would like to choose Credit Card as your payment type instead, please choose Cancel.");

                if (sel.options[sel.selectedIndex].value == "cc") {
                    var sel2 = document.getElementById("CardType");
                    if (sel2 != null && sel2.selectedIndex == 0) {
                        alert("Please choose your credit card type.");
                        return false;
                    }
                }
            }
        }

    </script>
</head>
<body onload='init();'>
    <!--#include file ="assets/header.asp"-->
    <div id="bodycontent" class="cart">
        <div class="container-fluid">
            <div class="container adjust-padding">
                <div class="row">
                    <div class="col-sm-12">
                        <h1>Payment Options</h1>
                        <%
        If Len(Session("ErrMsg")) > 1 Then
            Response.Write "<p class='error'><b>Payment was unsuccessful</b><br>" & Session("ErrMsg") & "</p>"
        End If
	Session("ErrMsg") = ""
	
	Select Case Request.QueryString("Error")
	    Case "Type" 
		Response.Write "<p class='error'>Please select payment type.</p>"
	    Case "Card"
		Response.Write "<p class='error'>Please select credit card type.</p>"
	    Case "Fill"
		Response.Write "<p class='error'>Please enter credit card information marked in red.</p>"
	End Select
                        %>
                        <p>
                            Enter your payment information here. This shopping cart section of the CHI website is secured with industry standard 128
                            bit encryption.
                        </p>
                    </div>
                </div>
            </div>
        </div>

        <form method="POST" action="paymentform.asp?Action=Do" name="form1" id="form1" onsubmit="return validateForm()">
            <div class="container-fluid white">
                <div class="container checkout">
                <!-- cart -->
                <div class="row">
                    <div class="col-sm-12">
                        <div class="title-line">
                            Order Summary
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
        IsVietnamWorkshop = False
          For t = 1 to Session("NumCartItems")
                  TotalItemCost = (Session("CartItemCost(" & t & ")") * Session("CartItemQuantity(" & t & ")"))
                If Not IsVietnamWorkshop And Session("CartItem(" & t & ")") = "W1018" Then
                    IsVietnamWorkshop = True
                End If
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

                    <%
    ' put in the credit card logo
    If Session("BillCountry") = "Canada" Then
                    %>
                    <div>
                        <img src="assets/3cards.gif" border="0">
                    </div>
                    <% Else %>
                    <div>
                        <img src="assets/2cards.gif" border="0">
                    </div>
                    <% End If %>

                    <!-- payment type -->
                    <div class="row">
                        <div class="col-sm-12">
                            <div class="title-line">
                                Payment Options
                            </div>
                            <table class="payment">
                                <colgroup>
                                    <col style="width: 22%">
                                    <col style="width: 78%">
                                </colgroup>
                                <tr>
                                    <td>Payment Type</td>
                                    <td>&nbsp;
                <select size="1" name="CardType2" id="CardType2" onchange="check_text()">
                    <option>Choose payment type...</option>
                    <%
		    
            ' for vietnam workshop 2018
            'If Not IsVietnamWorkshop Then
                    If Session("CardType") = "VISA" or Session("CardType") = "Mastercard" or Session("CardType") = "American Express" Or Session("CardType2") = "cc" Then
                        Response.Write "<option value='cc' selected>Credit Card</option>"
                    Else
                        Response.Write "<option value='cc'>Credit Card</option>"
                    End If
            'End If
		    ' for sa workshop in 2011, we do not show them the invoice option
		    'If Not (Len(Session("CartItem(1)")) > 3 And (Left(Session("CartItem(1)"), 4) = "W550" Or Left(Session("CartItem(1)"), 4) = "W552" Or Left(Session("CartItem(1)"), 4) = "W551" Or Left(Session("CartItem(1)"), 4) = "W578")) Then
		    
                        
            ' for vietnam workshop 2018
            If IsVietnamWorkshop Then
                            Response.Write "<option>Invoice</option>"
                            Response.Write "<option>Cash</option>"
			    Response.Write "<option>Wire transfer</option>"
                    
            ElseIf (Session("PayQuoteMode") = "Yes" And UCase(Left(Session("CurrentTransaction"), 1)) = "Q") Or (Session("PayQuoteMode") <> "Yes" And Session("CurrencyAbrev") = "CAD") Or (Session("PayQuoteMode") <> "Yes" And Session("ShipCountry") = "USA") Then
			    If Session("CardType") = "Invoice" Then
			        Response.Write "<option selected>Invoice</option>"
			    Else
			        Response.Write "<option>Invoice</option>"
			    End If
		    End If

         If Not IsVietnamWorkshop Then
		    ' only add wire transfer fee on US orders
		    If Session("CurrencyAbrev") = "US" And Session("ShipCountry") <> "USA" And Session("IsPosted") <> "Yes" Then
			If Session("CardType") = "Wire transfer" Then
			    Response.Write "<option selected>Wire transfer</option>"
			Else
			    Response.Write "<option>Wire transfer</option>"
			End If
		    End If
		 End If
                    %>
                </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Credit Card Type</td>
                                    <td>&nbsp;
                <select size="1" name="CardType" id="CardType">
                    <option>Choose card type...</option>
                    <%
                    If Session("CardType") = "VISA" Then
                        Response.Write "<option selected>VISA</option>"
                    Else
                        Response.Write "<option>VISA</option>"
                    End If
                    If Session("CardType") = "Mastercard" Then
                        Response.Write "<option selected>Mastercard</option>"
                    Else
                        Response.Write "<option>Mastercard</option>"
                    End If
		    If Session("CurrencyAbrev") = "CAD" Then
			If Session("CardType") = "American Express" Then
			    Response.Write "<option selected>American Express</option>"
			Else
			    Response.Write "<option>American Express</option>"
			End If
		    End If
		    
                    %>
                </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Card Number
                                    </td>
                                    <td>&nbsp;
                    <input type="text" name="CardNumber" size="40" value="<% Response.Write Session("CardNumber") %>" <% If (Session("CardType") = "VISA" or Session("CardType") = "Mastercard" or Session("CardType") = "American Express") And Session("CardNumber") = "" Then Response.Write "class='txtBackground'" %>>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Expiry Date (MM/YY or MMYY)
                                    </td>
                                    <td>&nbsp;
	  <input type="text" name="CardExpiry" size="40" value="<% Response.Write Session("CardExpiry") %>" <% If (Session("CardType") = "VISA" or Session("CardType") = "Mastercard" or Session("CardType") = "American Express") And Session("CardExpiry") = "" Then Response.Write "class='txtBackground'" %>>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Cardholder Name
                                    </td>
                                    <td>&nbsp;
	  <input type="text" name="CardName" size="40" value="<% Response.Write Session("CardName") %>" <% If (Session("CardType") = "VISA" or Session("CardType") = "Mastercard" or Session("CardType") = "American Express") And Session("CardName") = "" Then Response.Write "class='txtBackground'" %>>
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </div>
                    
                    <div class="row">
                        <div class="col-sm-12">
                           <div class="title-line">International Transaction Fees</div>
                            <p>Please note your credit card issuer may add a currency conversion or foreign transaction fee to your payment if you are outside of Canada. If you have a larger than expected charge on your statement, please contact your provider for information on what rates and fees may apply, as these are not controlled by or known to CHI.</p>
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
                                    <% If Session("PaymentForm") <> "Yes" Then %>
                                    <button type="submit" form="form1" class="btn btn-teal">
                                        NEXT
                                    <img class="arrow" src="assets/arrow-right.png" /></button>
                                    <% Else %>
                                    <button type="submit" form="form1" class="btn btn-teal">
                                        NEXT
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
