<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<%
'
' very important - should we abandon previous sessions before we proceed this step?
' it could mess things up if clients already got something in their cart - but how often would that happen?
'
	'load the transaction (if not loaded)
	If Len(Request.QueryString("Item")) < 7 Then
	    ' check item number
	    If InStr(Trim(Request.QueryString("Item")), " ") <> 0 Or InStr(Trim(Request.QueryString("Item")), "'") <> 0 Then
		Response.Redirect "http://www.chiwater.com/pay.asp?Error=Error"
	    End If 
	    
		Session("CurrentTransaction") = Trim(Request.QueryString("Item"))
		Session("QuoteName") = Trim(Request.QueryString("Name"))
		
		'Response.Redirect "loadtransaction.asp"
	
    	' set connection object here
	    set conn=Server.CreateObject("ADODB.Connection")
	    conn.Provider="Microsoft.Jet.OLEDB.4.0"
	    conn.Open transcationPath


		' try to catch all digits quote number before the switch
		Set objRec = Server.CreateObject ("ADODB.Recordset")	
		'objRec.Open "Transaction", transcationPath, 3, 3, 2
        objRec.Open "[Transaction]", conn
		objRec.Filter = "TransactionNumber = '" & Session("CurrentTransaction") & "'"
		' I know this is not the best way to open another reader in the db. not sure how to say this in vb - please modify
		
		If objRec.EOF And objRec.BOF Then
		    Set objRecTry = Server.CreateObject ("ADODB.Recordset")	
		    'objRecTry.Open "Transaction", transcationPath, 3, 3, 2
            objRecTry.Open "[Transaction]", conn
		    objRecTry.Filter = "TransactionNumber = 'Q" & Session("CurrentTransaction") & "'"		
		    
		    If objRecTry.EOF And objRecTry.BOF Then
			Response.Redirect "http://www.chiwater.com/pay.asp?Error=NotFound"
		    Else
			If objRecTry("Quote") = "Yes" Then Session("CurrentTransaction") = "Q" & Session("CurrentTransaction")
		    End If
		
		    objRecTry.Close
		    Set objRecTry = Nothing
		
		End If

		objRec.Close
		Set objRec = Nothing

		' load transaction here - we need the sessions for the next steps
		Set objRec = Server.CreateObject ("ADODB.Recordset")	
		objRec.Open "[Transaction]", transcationPath, 3, 3, 2
		objRec.Filter = "TransactionNumber = '" & Session("CurrentTransaction") & "'"
		
		If objRec.EOF And objRec.BOF Then
		    Response.Redirect "http://www.chiwater.com/pay.asp?Error=NotFound"
		Else
		    objRec.MoveFirst

		    ' check transactioncreated field to make sure this is not a double payment
		    If Not IsNull(ObjRec("TransactionCreated")) Then
			If ObjRec("TransactionCreated") >= ObjRec("OrderDate") Then
			    Response.Redirect "http://www.chiwater.com/pay.asp?Error=Paid"
			Else
			    Response.Redirect "http://www.chiwater.com/pay.asp?Error=Error"
			End If
		    End If
		    
		    Session("PayQuoteMode") = "Yes"

		    'Session("AllowEdit") = False
		    ' quote has expire date but order does not
		    isValid = False
		    'Session("AllowEdit") = True
		    If Not (Left(ObjRec("Items"), 4) = "C023" Or Left(ObjRec("Items"), 4) = "C323" Or Left(ObjRec("Items"), 4) = "C423") Or DateDiff("d", ObjRec("OrderDate"), CDate("February 26, 2014")) < 0 Then
			If ObjRec("Quote") = "Yes" Then 
			    If (DateDiff("d", DateAdd("m", 2, ObjRec("OrderDate")), Date) < 1) Then isValid = True
			Else
			    isValid = True
			    'Session("AllowEdit") = False
			End If
		    Else
			' this is a year 2012 conference quote - valid until February 22, 2012 
			If (DateDiff("d", CDate("February 26, 2014"), Date) > 1) And ObjRec("Quote") = "Yes"  Then
			'    isValid = True
			'Else
			    Response.Redirect "http://www.chiwater.com/pay.asp?Error=Expired"
                        Else
                            isValid = True
			End If
		    End If
		    
		    Session("QuoteQuery") = "PCSWMM"
		    
		    ' new rules: let people change their payment options for an order as long as it's not posted - in the same time, we are taking the risk that if this order
		    ' is being posted starting now until they submit their order. we might end up with the wire transfer fee (plus tax if app.) difference in our quickbook
		    If ObjRec("PostedDate") <> "" And ObjRec("PostedDate") >= ObjRec("OrderDate") Then Session("IsPosted") = "Yes"
		    
		    If isValid And (LCase(ObjRec("ClientLastName")) = LCase(Session("QuoteName"))) Then
			Session("OrderDate") = ObjRec("OrderDate")
                        Session("ProcessedDate") = ObjRec("ProcessedDate")
			'Session("PaidDate") = ObjRec("PaidDate")
                        Session("ContactID") = ObjRec("ContactID")
                        Session("OfficeID") = ObjRec("OfficeID")
                        Session("CompanyID") = ObjRec("CompanyID")
	
            Session("SourceOption")	= objRec("Source") 	

			'load client info ===========================
		
			Session("ClientAddress") = "Yes"
			Session("ClientPrefix") = ObjRec("ClientPrefix")
			Session("ClientFirstName") = ObjRec("ClientFirstName")
			Session("ClientLastName") = ObjRec("ClientLastName")
			Session("ClientName") =  Trim(Session("ClientPrefix") & " " & Trim(Session("ClientFirstName") & " " & Session("ClientLastName")))
			Session("ClientCompany") = ObjRec("ClientCompany")
			Session("ClientStreetAddress") = ObjRec("ClientStreetAddress")
			Session("ClientStreetAddress2") = ObjRec("ClientStreetAddress2")
			Session("ClientCity") = ObjRec("ClientCity")
			Session("ClientProvince") = ObjRec("ClientProvince")
			Session("ClientCountry") = ObjRec("ClientCountry")
			Session("ClientPostalCode") = ObjRec("ClientPostalCode")
			Session("ClientEmail") = ObjRec("ClientEmail")
			Session("ClientTelephone") = ObjRec("ClientTelephone")
			Session("ClientFax") = ObjRec("ClientFax")
			Session("ClientWebSite") = ObjRec("ClientWebSite")
			Session("ClientType") = ObjRec("ClientType")
		
			'load shipping info
		
			Session("ShipAddress") = "Yes"
			Session("ShipPrefix") = ObjRec("ShipPrefix")
			Session("ShipFirstName") = ObjRec("ShipFirstName")
			Session("ShipLastName") = ObjRec("ShipLastName")
			Session("ShipName") =  Trim(Session("ShipPrefix") & " " & Trim(Session("ShipFirstName") & " " & Session("ShipLastName")))
			Session("ShipCompany") = ObjRec("ShipCompany")
			Session("ShipStreetAddress") = ObjRec("ShipStreetAddress")
			Session("ShipStreetAddress2") = ObjRec("ShipStreetAddress2")
			Session("ShipCity") = ObjRec("ShipCity")
			Session("ShipProvince") = ObjRec("ShipProvince")
			Session("ShipCountry") = ObjRec("ShipCountry")
			Session("ShipPostalCode") = ObjRec("ShipPostalCode")
			Session("ShipEmail") = ObjRec("ShipEmail")
			Session("ShipTelephone") = ObjRec("ShipTelephone")
			Session("ShipFax") = ObjRec("ShipFax")
			
			'load billing info
		
			Session("BillAddress") = "Yes"
			Session("BillPrefix") = ObjRec("BillPrefix")
			Session("BillFirstName") = ObjRec("BillFirstName")
			Session("BillLastName") = ObjRec("BillLastName")
			Session("BillName") =  Trim(Session("BillPrefix") & " " & Trim(Session("BillFirstName") & " " & Session("BillLastName")))
			Session("BillCompany") = ObjRec("BillCompany")
			Session("BillStreetAddress") = ObjRec("BillStreetAddress")
			Session("BillStreetAddress2") = ObjRec("BillStreetAddress2")
			Session("BillCity") = ObjRec("BillCity")
			Session("BillProvince") = ObjRec("BillProvince")
			Session("BillCountry") = ObjRec("BillCountry")
			Session("BillPostalCode") = ObjRec("BillPostalCode")
			Session("BillEmail") = ObjRec("BillEmail")
			Session("BillTelephone") = ObjRec("BillTelephone")
			Session("BillFax") = ObjRec("BillFax")
			
			'load payment info ===================================
			
			Session("CardType") = ObjRec("CardType")
			' in pay mode, they should not use the payment information stored previously
			'Select Case Session("CardType")
			'	Case "VISA", "Mastercard", "American Express", "Wire transfer"
			'	Case Else
			'		Session("CardType") = ""
			'End Select
			'Session("CardNumber") = ObjRec("CardNumber")
			'Session("CardExpiry") = ObjRec("CardExpiry")
			'Session("CardName") = ObjRec("CardName")
			'Session("CardAuthorizationNo") = ObjRec("CardAuthorizationNo")
		
			'load currencies ======================================
			
			Session("Currency") = ObjRec("Currency")
			Session("CurrencyAbrev") = ObjRec("CurrencyAbrev")
			'
			' todo for the future if the currency rate is different, change below code
			'
			'Session("CurrencyExchange") = 0 ' assign default value
			'Session("CurrencyExchange") = ObjRec("CurrencyExchange")
			'If IsNull(Session("CurrencyExchange")) Then Session("CurrencyExchange") = 0
			
			Session("ShippingCost") = ObjRec("ShippingCost")
			If IsNull(Session("ShippingCost")) Then Session("ShippingCost") = 0
		
			Session("HSTCost") = ObjRec("HSTCost")
			If IsNull(Session("HSTCost")) Then Session("HSTCost") = 0
			Session("GSTCost") = ObjRec("GSTCost")
			If IsNull(Session("GSTCost")) Then Session("GSTCost") = 0
			Session("PSTCost") = ObjRec("PSTCost")
			If IsNull(Session("PSTCost")) Then Session("PSTCost") = 0
		
			Session("SubTotalCost") = ObjRec("SubTotalCost")
			If IsNull(Session("SubTotalCost")) Then Session("SubTotalCost") = 0
			Session("TotalCost") = ObjRec("TotalCost")
			If IsNull(Session("TotalCost")) Then Session("TotalCost") = 0
			
			Session("ShippingRequired") = ObjRec("ShippingRequired")
			'Session("NoGST") = ObjRec("NoGST")
			'Session("NoPST") = ObjRec("NoPST")
			'Session("NoHST") = ObjRec("NoHST")
			
			Session("ClientNotes") = ObjRec("ClientNotes")
                        Session("PurchaseOrder") = ObjRec("PurchaseOrder")
		
			'load items ===========================================
			Items = ObjRec("Items")
			Dim Item
			Dim ItemDescription
			Dim ItemCost
			Dim ItemQuantity
			n=1
		
			If Items = "" Then
					Session("NumCartItems") = 0
			Else
				For t = 1 to 1000
					'get item
					p = Instr(n, Items, "/")
					If p = 0 Then Exit For
					Item = Mid(Items, n, p-n)
					n = p + 1
					'get description
					p = Instr(n, Items, "/")
					If p = 0 Then Exit For
					ItemDescription= Mid(Items, n, p-n)
					n = p + 1
					'get cost
					p = Instr(n, Items, "/")
					If p = 0 Then Exit For
					ItemCost= Mid(Items, n, p-n)
					n = p + 1
					'get number
					p = Instr(n, Items, "/")
					If p = 0 Then Exit For
					ItemQuantity = Mid(Items, n, p-n)
					n = p + 1
					'create session item
					Session("NumCartItems") = t
					Session("CartItem(" & t & ")") = Item
					Session("CartItemDescription(" & t & ")") = ItemDescription

If Left(Session("CurrentTransaction"), 1) = "Q" And DateDiff("d", Session("OrderDate"), CDate(#4/18/22#)) < 0 Then
    PriceChange = 0
    If Item = "S222" And ItemCost = 1440 Then 
        PriceChange = 240
    ElseIf Item = "S220" And ItemCost = 2160 Then 
        PriceChange = 160
    End If
    If PriceChange > 0 Then
        ItemCost = ItemCost + PriceChange
        Session("SubTotalCost") = Session("SubTotalCost") + PriceChange
        	If Session("HSTCost") > 0 Then 
                taxRate = 0.13
                If Session("ShipProvince") = "PE" Then taxRate = 0.15
                Session("TotalCost") = Session("TotalCost") - Session("HSTCost")
                Session("HSTCost") = Session("HSTCost") + PriceChange * taxRate
                Session("TotalCost") = Session("TotalCost") + Session("HSTCost") + PriceChange
            End If
			If Session("GSTCost") > 0 Then 
                Session("TotalCost") = Session("TotalCost") - Session("GSTCost")
                Session("GSTCost") = Session("GSTCost") + PriceChange * 0.05
                Session("TotalCost") = Session("TotalCost") + Session("GSTCost") + PriceChange
            End If
    End If
    ' for us orders
    If Session("CurrencyAbrev") = "US" Then Session("TotalCost") = Session("TotalCost") + PriceChange 
End If
					Session("CartItemCost(" & t & ")") = ItemCost
					Session("CartItemQuantity(" & t & ")") = ItemQuantity
					
					If LCase(Left(Session("CartItem(" & t & ")"), 1)) = "w" Then
					    Session("QuoteQuery") = "Workshops"
					End If 
					
					If n > len(Items) Then Exit For
				Next
			End If
		    Else
			If Not (DateDiff("d", DateAdd("m", 2, ObjRec("OrderDate")), Date) < 1) Then Response.Redirect "http://www.chiwater.com/pay.asp?Error=Expired"
			If Not (LCase(ObjRec("ClientLastName")) = LCase(Session("QuoteName"))) Then Response.Redirect "http://www.chiwater.com/pay.asp?Error=Name"
		    End If
		End If   
		objRec.Close
		Set objRec = Nothing
        conn.Close
	    Set conn = Nothing
	Else
	    Response.Redirect "http://www.chiwater.com/pay.asp?Error=Error"
	End If
	
	Response.Redirect("checkout.asp")
	
%>








 