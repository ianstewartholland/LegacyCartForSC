﻿<%
	Session.Timeout = 60
    ' workshop or software cart check for the quotes
    If IsNull(Session("QuoteQuery")) Then
        Session("QuoteQuery") = "PCSWMM"
    End If
    
    'update quantity
    If Request.QueryString("Action") = "UpdateCart" Then

	' for the advanced workshops - check R999X
	hasWBooks = False
	pos = 0
    
    ' check SA 2015, cannot register for just day 1 or 2 so if day 1 item is removed, day 2 shall also be removed
    Do
		For t = 1 to Session("NumCartItems")
            If Len(Session("CartItem(" & t & ")")) > 4 And (Right(Session("CartItem(" & t & ")"), 1) <> "C") And (Left(Session("CartItem(" & t & ")"), 4) = "W845" Or Left(Session("CartItem(" & t & ")"), 4) = "W846" Or Left(Session("CartItem(" & t & ")"), 4) = "W847") Then

                ' if day 1 is removed
			    If Request.Form("CartItemQuantity" & t) <=0 And Right(Session("CartItem(" & t & ")"), 1) = "A" Then
				    Session("NumCartItems") = Session("NumCartItems") - 1
				    For n = t to Session("NumCartItems")
					    Session("CartItem(" & n & ")") = Session("CartItem(" & n + 1 & ")")
					    Session("CartItemDescription(" & n & ")") = Session("CartItemDescription(" & n + 1 & ")")
					    Session("CartItemCost(" & n & ")") = Session("CartItemCost(" & n + 1 & ")")
					    Session("CartItemQuantity(" & n & ")") = Session("CartItemQuantity(" & n + 1 & ")")
				    Next
			    End If

                ' if day 2 is removed
			    If Request.Form("CartItemQuantity" & t) <=0 And Right(Session("CartItem(" & t & ")"), 1) = "B" Then

				    Session("NumCartItems") = Session("NumCartItems") - 1
				    For n = t to Session("NumCartItems")
					    Session("CartItem(" & n & ")") = Session("CartItem(" & n + 1 & ")")
					    Session("CartItemDescription(" & n & ")") = Session("CartItemDescription(" & n + 1 & ")")
					    Session("CartItemCost(" & n & ")") = Session("CartItemCost(" & n + 1 & ")")
					    Session("CartItemQuantity(" & n & ")") = Session("CartItemQuantity(" & n + 1 & ")")
				    Next
    				Session("NumCartItems") = Session("NumCartItems") - 1
				    For n = t - 1 to Session("NumCartItems")
					    Session("CartItem(" & n & ")") = Session("CartItem(" & n + 1 & ")")
					    Session("CartItemDescription(" & n & ")") = Session("CartItemDescription(" & n + 1 & ")")
					    Session("CartItemCost(" & n & ")") = Session("CartItemCost(" & n + 1 & ")")
					    Session("CartItemQuantity(" & n & ")") = Session("CartItemQuantity(" & n + 1 & ")")
				    Next
			    End If
            End If
		Next
		If t > Session("NumCartItems") Then Exit Do
	Loop 

        
        ' check enterprise items S236, S237
        hasEnterprisePerUser = False
        NumBase = 0

	For t = 1 to Session("NumCartItems")
            If Session("CartItem(" & t & ")") = "S236" Then
                NumBase = Request.Form("CartItemQuantity" & t)
            End If
            
            If Session("CartItem(" & t & ")") = "S237" And Request.Form("CartItemQuantity" & t) > 2 Then hasEnterprisePerUser = True
	Next
        
        If NumBase = 1 And Not hasEnterprisePerUser Then Session("Error") = "<p class = 'err'>Enterprise license requires a minimum of 3 users.</p>"
        
        If hasEnterprisePerUser And NumBase <> 1 Then Session("Error") = "<p class = 'err'>Enterprise license requires one PCSWMM Enterprise subscription base item S236.</p>"
        
        
        If Len(Session("Error")) > 1 Then Response.Redirect "cart.asp"
        
	'update all items
	For t = 1 to Session("NumCartItems")
		If Request.Form("CartItemQuantity" & t) < 0 Then Response.Redirect "cart.asp"
		Session("CartItemQuantity(" & t & ")")= Request.Form("CartItemQuantity" & t)
		'Session("TotalNumItems") = Session("TotalNumItems") + Session("CartItemQuantity(" & t & ")")
		If Session("CartItem(" & t & ")") = "R999X" Then
		    hasWBooks = True
		    pos = t
		End If
	Next
	
	' check item quantity, whether the discount and the original quantity matches
	For t = 1 To Session("NumCartItems")
	    'find discount item
	    If Right(Session("CartItem(" & t & ")"), 1) = "X" Or Right(Session("CartItem(" & t & ")"), 1) = "G" Then
		item = Left(Session("CartItem(" & t & ")"), 4)
		quan = Session("CartItemQuantity(" & t & ")")
		found = False
		For i = 1 To Session("NumCartItems")
		    If i <> t And InStr(Session("CartItem(" & i & ")"), item) > 0 Then
			found = true
			'check regular item quantity
			If Session("CartItemQuantity(" & i & ")") < quan Then
			    Session("CartItemQuantity(" & t & ")") =  Session("CartItemQuantity(" & i & ")")
			    Exit For
			End If
		    End If
		Next
		' remove item if not found
		If Not found Then
		    Session("CartItemQuantity(" & t & ")") = 0
		End If
	    End If
	Next
	
	If hasWBooks Then
	    ' check books and reset discount 
	    bookNum = 0
	    book = ""
	    t1 = 0
	    t2 = 0
	    For t = 1 To Session("NumCartItems")
		If Session("CartItem(" & t & ")") = "R184" And Session("CartItemQuantity(" & t & ")") > 0 Then
		    bookNum = bookNum + 1
		    book = "R184"
		    t1 = t
		End If
		If Session("CartItem(" & t & ")") = "R242" And Session("CartItemQuantity(" & t & ")") > 0 Then
		    bookNum = bookNum + 1
		    book = "R242"
		    t2 = t
		End If
	    Next
   
	    If bookNum = 0 Then
		Session("CartItemQuantity(" & pos & ")") = 0
	    ElseIf bookNum = 1 Then
		If Len(book) > 1 Then
		    If book = "R184" Then
			Session("CartItemCost(" & pos & ")") = - Session("CartItemCost(" & t1 & ")") * Session("CartItemQuantity(" & t1 & ")") * 0.2
		    Else
			Session("CartItemCost(" & pos & ")") = - Session("CartItemCost(" & t2 & ")") * Session("CartItemQuantity(" & t2 & ")") * 0.2
		    End If
		    Session("CartItemQuantity(" & pos & ")") = 1
		End If
	    ElseIf bookNum = 2 Then
		Session("CartItemCost(" & pos & ")") = - (Session("CartItemCost(" & t1 & ")") * Session("CartItemQuantity(" & t1 & ")") + Session("CartItemCost(" & t2 & ")") * Session("CartItemQuantity(" & t2 & ")")) * 0.2
		Session("CartItemQuantity(" & pos & ")") = 1
	    Else
		' should not happen
	    End If
	End If 

	'remove any items with zero quantity
	Do
		For t = 1 to Session("NumCartItems")
			If Session("CartItemQuantity(" & t & ")") = 0 Then
				Session("NumCartItems") = Session("NumCartItems") - 1
				For n = t to Session("NumCartItems")
					Session("CartItem(" & n & ")") = Session("CartItem(" & n + 1 & ")")
					Session("CartItemDescription(" & n & ")") = Session("CartItemDescription(" & n + 1 & ")")
					Session("CartItemCost(" & n & ")") = Session("CartItemCost(" & n + 1 & ")")
					Session("CartItemQuantity(" & n & ")") = Session("CartItemQuantity(" & n + 1 & ")")
				Next
				Exit For
			End If
		Next
		If t > Session("NumCartItems") Then Exit Do
	Loop 

	Response.Redirect "cart.asp"
    End If
    
    'need to update the number of cart items in the front end server (to display a link in the header)
'    If Session("UpdateCartStatus") = False Then
'        Session("UpdateCartStatus") = True
'        Response.redirect "http://www.chiwater.com/updatecartstatus.asp?numitems=" & Session'("NumCartItems")
'    Else
'        Session("UpdateCartStatus") = False
'    End If
    
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cart - CHI Shopping Cart</title>
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
                        <h1>Shopping Cart</h1>
                        <%            
  	If Session("NumCartItems") > 0 Then
        
        'remove duplicate free trials
        FreeTrial = 0
	Do
	    For t = 1 to Session("NumCartItems")
			If Session("CartItem(" & t & ")") = "S212" Then
                            FreeTrial = FreeTrial + 1
                        End If
                        If FreeTrial > 1 Then
			    Session("NumCartItems") = Session("NumCartItems") - 1
                            For n = t to Session("NumCartItems")
                                Session("CartItem(" & n & ")") = Session("CartItem(" & n + 1 & ")")
				Session("CartItemDescription(" & n & ")") = Session("CartItemDescription(" & n + 1 & ")")
				Session("CartItemCost(" & n & ")") = Session("CartItemCost(" & n + 1 & ")")
				Session("CartItemQuantity(" & n & ")") = Session("CartItemQuantity(" & n + 1 & ")")
			    Next
			    Exit For
			End If
	    Next
	    If t > Session("NumCartItems") Then Exit Do
	Loop
        
        ' display enterprise error
        If Len(Session("Error")) > 0 Then
            Response.Write Session("Error") 
        End If
        
        Session("Error") = ""
        
  		If Session("NumCartItems") = 1 Then
                        %>
                        <p>
                            You have <% Response.Write Session("NumCartItems") %> item in your cart.
	<%
		Else
    %>
                        <p>
                            You have <% Response.Write Session("NumCartItems") %> items in your cart.
	<%
		End If
    %>
        To proceed with your order, click on the <i>Check Out</i> button below. You may change the quantity of an item, or put zero in the quantity
                            to remove an item. To save your changes, click on the
        <i>Update Quantity</i> button.
                        </p>

                        <%
        'check for free trial order
        FreeTrialOrder = ""
        If Session("NumCartItems") = 1 Then
            If Session("CartItem(1)") = "S212" Then
                        %>
                        <p>
                            <b>Free Trial Licenses:</b> Valid client information including shipping address <b>is required</b>. Only serious inquiries
                            with valid contact information will be accepted.
                Once the client information has been verified by our staff, you will be emailed a download link.
                        </p>
                        <%
            End If
        End If
    Else
                        %>
                        <p>You have no items in your cart.</p>
                        <%
    End If
                        %>
                    </div>
                </div>
            </div>
        </div>
        <% If Session("PayQuoteMode") <> "Yes" Then %>
        <form method="POST" action="cart.asp?Action=UpdateCart" name="cartform" id="cartform">
            <% End If %>
            <div class="container-fluid white">
                <div class="container checkout">
                    <div class="row">
                        <div class="col-sm-12">
                            <div class="table-row shopping-cart">
                                <div class="row title-line">
                                    <div class="col-md-1 col-sm-2 col-xs-2">
                                        ITEM
                                    </div>
                                    <div class="col-md-7 col-sm-4 col-xs-10">
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
          dim TotalCost
          dim TotalItemCost
          TotalCost = 0
	  If Session("PayQuoteMode") <> "Yes" Then Session("ShippingRequired") = ""
          For t = 1 to Session("NumCartItems")
                TotalItemCost = (Session("CartItemCost(" & t & ")") * Session("CartItemQuantity(" & t & ")"))
	    If Session("PayQuoteMode") <> "Yes" Then
		If LCase(Left(Session("CartItem(" & t & ")"), 1)) = "r" And Len(Session("CartItem(" & t & ")")) < 5 And Session("PayQuoteMode") <> "Yes" Then
		    Session("ShippingRequired") = "Yes"
		Else
		    Session("ShippingCost") = 0
		End If
	    End If
	    
	    ' set quote control here
	    If LCase(Left(Session("CartItem(" & t & ")"), 1)) = "w" Then
		Session("QuoteQuery") = "Workshops"
	    End If 
                                %>
                                <div class="row">
                                    <div class="col-md-1 col-sm-2 col-xs-2">
                                        <% response.write Session("CartItem(" & t & ")") %>
                                    </div>
                                    <div class="col-md-7 col-sm-4 col-xs-10">
                                        <% response.write Session("CartItemDescription(" & t & ")") %>
                                    </div>
                                    <div class="col-sm-2 right hidden-xs">
                                        <% response.write FormatCurrency(Session("CartItemCost(" & t & ")"),2) %>
                                    </div>
                                    <div class="col-md-1 col-sm-2 col-xs-12 right">
                                        <div class="visible-xs-inline"><% response.write FormatCurrency(Session("CartItemCost(" & t & ")"),2) %> <b class="purple">*</b> </div>
                                        <% If Session("PayQuoteMode") <> "Yes" Then %>
                                        <input type="text" style="text-align: right" name="CartItemQuantity<% response.write t %>" size="5" value="<% response.write Session("CartItemQuantity(" & t & ")") %>">
                                        <% Else %>
                                        <% response.write Session("CartItemQuantity(" & t & ")") %>
                                        <% End If %>
                                    </div>
                                    <div class="col-md-1 col-sm-2 hidden-xs">
                                        <% response.write FormatCurrency(TotalItemCost,2) %>
                                    </div>
                                </div>

                                <%
            TotalCost = TotalCost + TotalItemCost
            Next
	    ' assign temp total cost for page redirection on the checkout page
	    Session("SubTotalCost") = TotalCost
                                %>
                            </div>
                            <div class="row teal summary-line">
                                <div class="col-md-11 col-sm-10 hidden-xs">
                                    <b>Subtotal:</b>
                                </div>
                                <div class="col-md-1 col-sm-2 col-xs-12">
                                    <div class="visible-xs-inline"><b>Subtotal:</b></div>
                                    <b><% response.write FormatCurrency(TotalCost,2) %></b>
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
                            <% If Session("NumCartItems") > 0 Then %>
                            <div class="row button-row">
                                <div class="col-xs-12 right">
                                    <% If Session("PayQuoteMode") <> "Yes" Then %>
                                    <a href="checkout.asp?quote=yes" class="btn btn-teal">VIEW QUOTE
                                    <img class="arrow" src="assets/arrow-right.png" /></a>
                                    <% End If %>
                                    <button type="submit" form="cartform" class="btn btn-teal">
                                        UPDATE SUBTOTAL
                                    <img class="arrow" src="assets/arrow-right.png" /></button>

                                    <a href="checkout.asp" class="btn btn-blue">CHECK OUT
                                    <img class="arrow" src="assets/arrow-right.png" /></a>
                                </div>
                            </div>
                            <% End If %>
                            <div class="more-info">
                                <p class="title">NEED MORE INFORMATION?</p>
                                <p>
                                    If you have any questions, please email <a href="mailto:info@chiwater.com">info@chiwater.com</a>
                                    or call us at 1-519-767-0197
                                    or 1-888-972-7966.
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <% If Session("PayQuoteMode") <> "Yes" Then %>
        </form>
        <% End If %>
    </div>

    <!--#include file ="assets/footer.asp"-->
</body>
</html>
