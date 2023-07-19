<%
	'check to ensure CartItems exist
	If Session("NumCartItems") > 0 Then
		'everything ok
	Else
		Response.Redirect("cart.asp")
	End If

%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Shopping Cart Check Out - CHI Shopping Cart</title>
    <link rel="shortcut icon" href="assets/chilogo.ico" />
    <link href='https://fonts.googleapis.com/css?family=PT+Serif|Fjalla+One|Open+Sans:300italic,400italic,600italic,300,400,600,700'
        rel='stylesheet' type='text/css'>
    <link href="assets/site.min.css" rel="stylesheet" />
    <link href="assets/style.min.css" rel="stylesheet" />
    <script src="assets/startup.service.js"></script>
    <script src="assets/cart.js"></script>
</head>
<body>
<!-- Google Code for Checkout Conversion Page  May 4, 2011 -->
<script type="text/javascript">
    /* <![CDATA[ */
    var google_conversion_id = 1066628046;
    var google_conversion_language = "en";
    var google_conversion_format = "2";
    var google_conversion_color = "ffffff";
    var google_conversion_label = "8EJ7CIq7ywEQzufN_AM";
    var google_conversion_value = 0;
    /* ]]> */
</script>
<script type="text/javascript" src="http://www.googleadservices.com/pagead/conversion.js">
</script>
<noscript>
<div style="display:inline;">
<img height="1" width="1" style="border-style:none;" alt="" src="http://www.googleadservices.com/pagead/conversion/1066628046/?label=8EJ7CIq7ywEQzufN_AM&amp;guid=ON&amp;script=0"/>
</div>
</noscript>
    <!--#include file ="assets/header.asp"-->
    <div id="bodycontent" class="cart">
        <div class="container-fluid">
            <div class="container adjust-padding">
                <div class="row">
                    <div class="col-sm-12">
                        <h1>Shopping Cart Check Out</h1>
                        <p>
                            Your order was <b>successful</b><% If Session("CardType") = "Credit Card" Then Response.Write " and your credit card has been billed" %>.
                            Your transaction number is: <b><% Response.Write Session("CurrentTransaction") %></b>. Please print this page for your records.
                        </p>
                        <p>
                            An order confirmation will be emailed to <b><% Response.Write Session("ClientEmail") %></b>
                            when your order is processed during regular business hours, Monday to Friday from 8 a.m. to 4:30 p.m. EST. <% If Session("isTrial") Then Response.Write " You will receive a separate trial license activation email once your order has been processed." %>
                        </p>
                        <p>Thanks for supporting our business.</p>
                    </div>
                </div>
            </div>
        </div>

        <div class="container-fluid white">
            <div class="container checkout">
                <div class="row">
                    <div class="col-sm-12 ">
                        <div class="title-line">
                            Client Information
                        </div>
                        <div>

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
                        </div>
                    </div>
                </div>

                <% If Session("SubTotalCost") > 0 Then %>
                <!-- billing - no shipping -->
                <div class="row">
                    <div class="col-sm-12 ">
                        <div class="title-line">
                            Billing Address
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

                <!-- payment method -->
                <div class="row">
                    <div class="col-sm-12 ">
                        <div class="title-line">
                            Payment Method
                        </div>
                        <div>
                            <p>
                                <%

	Select Case Session("CardType")
	    Case "Invoice"
		Response.Write "Invoice. Payment due immediately upon receipt."
	    Case "VISA", "Mastercard", "American Express"
    		Response.Write "Card type: " & Session("CardType") & "<br>"
		Response.write "Card number: " & Session("CardNumber") & "<br>"
    		Response.Write "Cardholder name: " & Session("CardName") & ""
	    Case "Wire transfer"
                                %>
			Wire transfer in <b>US dollars</b> to:
                                <br>
                                <br>
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
                                <br>
                                <br>
                                A fee of $20 has been added to cover bank charges. Please <b>email a PDF proof of payment</b> to <a href="mailto:info@chiwater.com?subject=Wire transfer">
                                    info@chiwater.com</a>

                            </p>
                            <p>
                                <b>Please use your transaction
number <% Response.Write Session("CurrentTransaction") %> as reference when you make payment to allow us to allocate
payment correctly.</b>
                                <% End Select %>
                            </p>
                        </div>
                    </div>
                </div>
                <% End If %>

                <!-- cart -->
                <div class="row">
                    <div class="col-sm-12">
                        <div class="title-line">
                            Shopping Cart
                        </div>
                        <table class="checkout">
                            <colgroup>
                                <col style="width: 10%">
                                <col style="width: 58%">
                                <col style="width: 10%">
                                <col style="width: 14%">
                                <col style="width: 8%">
                            </colgroup>
                            <tr>
                                <td>ITEM
                                </td>
                                <td>DESCRIPTION
                                </td>
                                <td>UNIT PRICE
                                </td>
                                <td>QUANTITY
                                </td>
                                <td>AMOUNT
                                </td>
                            </tr>
                            <%
          For t = 1 to Session("NumCartItems")
                  TotalItemCost = (Session("CartItemCost(" & t & ")") * Session("CartItemQuantity(" & t & ")"))
                toPrint = False
                If t = Session("NumCartItems") Then toPrint = True
                            %>
                            <tr>
                                <td>
                                    <% response.write Session("CartItem(" & t & ")") %>
                                </td>
                                <td <% If toPrint Then Response.Write "style='border-bottom:none;'" %>>
                                    <% response.write Session("CartItemDescription(" & t & ")") %>
                                </td>
                                <td <% If toPrint Then Response.Write "style='border-bottom:none;'" %>>
                                    <% response.write FormatCurrency(Session("CartItemCost(" & t & ")"),2) %>
                                </td>
                                <td <% If toPrint Then Response.Write "style='border-bottom:none;'" %>>
                                    <% response.write Session("CartItemQuantity(" & t & ")") %>
                                </td>
                                <td>
                                    <% response.write FormatCurrency(TotalItemCost,2) %>
                                </td>
                            </tr>
                            <% Next %>
                            <tr>
                                <td colspan="4" class="no-border reduce-padding">Subtotal:
                                </td>
                                <td class="no-border reduce-padding">
                                    <% response.write FormatCurrency(Session("SubTotalCost"),2) %>
                            </tr>
                            <% If Session("HSTCost") > 0 Then %>
                            <tr>
                                <td colspan="4" class="no-border reduce-padding">HST (#119880979):</td>
                                <td class="no-border reduce-padding"><% response.write FormatCurrency(Session("HSTCost"),2) %></td>
                            </tr>
                            <%
                End If
  		If Session("GSTCost") > 0 Then
                            %>
                            <tr>
                                <td colspan="4" class="no-border reduce-padding">GST (#119880979):</td>
                                <td class="no-border reduce-padding"><% response.write FormatCurrency(Session("GSTCost"),2) %></td>
                            </tr>
                            <%
  		End If
                            %>
                            <tr>
                                <td colspan="4" class="no-border">Total (<% Response.Write Session("CurrencyAbrev") %>$):</td>
                                <td class="no-border"><% response.write FormatCurrency(Session("TotalCost"),2) %></td>
                            </tr>
                        </table>
                    </div>
                </div>

            </div>
        </div>
        <div class="container-fluid no-print">
            <div class="container">
                <div class="row">
                    <div class="col-sm-12">
                        <div class="row button-row">
                            <div class="col-xs-12 right">
                                <button class="btn btn-teal" onclick="window.print()">PRINT THIS PAGE</button>
                            </div>
                        </div>
                        <p>
                            <b>Note</b>: You may need to remove the headers and footers that your browser includes
        on the printed page. This can typically be done through the print setup dialog for
        your browser. For best print quality, please use Google Chrome, IE 8+ or Firefox.
                        </p>
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
        <%
	Session.Abandon
        %>
    </div>

    <!--#include file ="assets/footer.asp"-->
</body>
</html>
