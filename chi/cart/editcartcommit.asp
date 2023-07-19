<%

	dim t
	dim n
	
	'reset the total number of items count
	Session("TotalNumItems") = 0
	
	'update all items
	For t = 1 to Session("NumCartItems")
	    If Request.Form("CartItemQuantity" & t) < 0 Then Response.Redirect "cart.asp"
		Session("CartItemQuantity(" & t & ")")= Request.Form("CartItemQuantity" & t)
		Session("TotalNumItems") = Session("TotalNumItems") + Session("CartItemQuantity(" & t & ")")
	Next
        
        's check item quantity, whether the discount and the original quantity matches
        For t = 1 To Session("NumCartItems")
            'find discount item
            If Right(Session("CartItem(" & t & ")"), 1) = "X" Then
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

	'set flags for recalculating
	'Session("SubTotalCalculated") = ""
	'Session("ShippingCalculated") = ""
	'Session("ShippingRequired") = ""
	'Session("TaxCalculated") = ""
	'Session("TotalCost") = 0

	'if backend user then mark changes
	'If Session("PasswordOK") = "Yes" and Session("CurrentTransaction") <> 0 then Session("Edited") = True
	
	Response.Redirect "cart.asp"

%>








 