<%
	'check to ensure CartItems exist
	If Session("NumCartItems") > 0 Then
		'everything ok
		If Session("SubTotalCost") <= 0 Then Response.Redirect("checkout.asp")
	Else
		Response.Redirect("cart.asp")
	End If
	
	If Request.QueryString("Action") = "Do" Then
	    Session("PurchaseOrder") = Request.Form("PurchaseOrder")
	    Response.Redirect "checkout.asp"
	End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Purchase Order - CHI Shopping Cart</title>
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
                        <h1>Purchase Order</h1>
                    </div>
                </div>
            </div>
        </div>
        <form method="POST" action="purchaseorder.asp?Action=Do" name="form" id="form">
            <div class="container-fluid white">
                <div class="container checkout">
                    <div class="row">
                        <div class="col-sm-12">
                            <div class="title-line">
                                Enter your own purchase order number (if any)
                            </div>
                            <table class="payment">
                                <colgroup>
                                    <col style="width: 22%">
                                    <col style="width: 78%">
                                </colgroup>
                                <tr>
                                    <td style="border-bottom:none !important;">Purchase order number</td>
                                    <td style="border-bottom:none !important;">&nbsp;<input type="text" name="PurchaseOrder" size="20" value="<% Response.Write Session("PurchaseOrder") %>">
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </form>
        
            <div class="container-fluid">
                <div class="container">
                    <div class="row">
                        <div class="col-sm-12">
                            <div class="row button-row">
                                <div class="col-xs-12 right">
                                    <button type="submit" form="form" class="btn btn-teal">
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
    </div>

    <!--#include file ="assets/footer.asp"-->
</body>
</html>
