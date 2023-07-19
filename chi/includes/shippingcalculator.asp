<!--#include file ="purolator.asp"-->
<!--#include file ="printxml.asp"-->
<!--#include file ="WriteLogEntry.asp"-->
<!--#include file ="dbpaths.inc"-->
<%
'
' changes to the shipping calculator - no longer calculate shipping for web training in Canada or USA
'
' function to calculate the shipping cost
Function CalculateShipping (Items, _
			    ClientCountry, _
			    ClientProvince, _
			    ClientCity, _
			    ClientPostalCode, _
			    CurrencyStr)
  	NumBooks = 0
	ShipWorkbook = False	
	' parse items
	Dim Item
	'Dim ItemDescription
	'Dim ItemCost
	Dim ItemQuantity
	n=1
	
	'define array to store the item information
	Dim CartItem()
	'Dim CartItemDescription()
	'Dim CartItemCost()
	Dim CartItemQuantity()
	If Items = "" Then
			NumCartItems = 0
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
			NumCartItems = t
			
			redim preserve CartItem(t)
			CartItem(t) = Item
			
			'redim preserve CartItemDescription(t)
			'CartItemDescription(t) = ItemDescription
			'
			'redim preserve CartItemCost(t)
			'CartItemCost(t) = ItemCost
			'
			redim preserve CartItemQuantity(t)
			CartItemQuantity(t) = ItemQuantity

			If n > len(Items) Then Exit For
		Next
	End If
	
	isWebTraining = False
	' new rules: calculate shipping when it has a book (as in publications or in software or in web training with an international address)
	For t = 1 to UBound(CartItem)
	    If Left(CartItem(t), 1) = "W" And Len(CartItem(t)) < 5 Then
		' you can look up the item description but to be safe I will look up in the workshop database to see whether this is a web training
		Set objRecWk = Server.CreateObject ("ADODB.Recordset")	
		objRecWk.Open "Workshops", workshopPath, 3, 3, 2
		objRecWk.Filter = "WorkshopID = '" & CartItem(t) & "'"
		objRecWk.MoveFirst
		If InStr(LCase(objRecWk("Country")), "web training") > 0 And ClientCountry <> "Canada" And ClientCountry <> "USA" Then isWebTraining = True
		objRecWk.Close
		Set objRecWk = Nothing
	    End If
	Next

	'
	' remember to add books for the web training here - since we combine the cost for books and training together for most cases
	'
	If isWebTraining Then
	    redim preserve CartItem(NumCartItems + 1)
	    CartItem(NumCartItems + 1) = "R184"
	    redim preserve CartItemQuantity(NumCartItems + 1)
	    CartItemQuantity(NumCartItems + 1) = 1
			
	    redim preserve CartItem(NumCartItems + 2)
	    CartItem(NumCartItems + 2) = "R231"
	    redim preserve CartItemQuantity(NumCartItems + 2)
	    CartItemQuantity(NumCartItems + 2) = 1
			
	    redim preserve CartItem(NumCartItems + 3)
	    CartItem(NumCartItems + 3) = "R242"
	    redim preserve CartItemQuantity(NumCartItems + 3)
	    CartItemQuantity(NumCartItems + 3) = 1
			
	    NumCartItems = NumCartItems + 3
	End If 

	Set BookList = Server.CreateObject("Scripting.Dictionary")
  	'count the number of books and the number of software orders
  	For n = 1 to NumCartItems
		If Left(CartItem(n), 1) = "R" And Len(CartItem(n)) < 5 Then
                    'Dim bookID
                    bookID = CartItem(n)
                    NumBooks = NumBooks + CartItemQuantity(n)
WriteLogEntry("Has " & CartItemQuantity(n) & bookID & "book(s)")
                    'If BookList.Exists(bookID) Then
                    '    BookList.Item(bookID) = BookList.Item(bookID) + 1
                    'Else
                    '    BookList.Add bookID, 1
                    'End If
		    
                    If BookList.Exists(bookID) Then
                        BookList.Item(bookID) = BookList.Item(bookID) + CartItemQuantity(n)
                    Else
                        BookList.Add bookID, CartItemQuantity(n)
                    End If
		    
		End If
  	Next

WriteLogEntry( "Total book number: " & 	NumBooks)
	
	's re-write this when have time
  	'check whether shipping required has been determined
'  	Select Case ShippingRequired
'  		Case "Yes", "No" 
'  			'do nothing
'  		Case Else
'  			'determine if shipping is required
'  			ShippingRequired = "Yes"
'		  	'if we are meeting face to face then assume nothing needs to be shipped
'		  	For n = 1 to NumCartItems
'				Select Case Left(CartItem(n), 1)
'					Case "W"
'						If Lcase(Left(CartItemDescription(n), 10)) <> "self-paced" Then 
'							ShippingRequired = "No"
'						Else
'							ShipWorkbook = True
'						End If
'					Case "I", "E", "C", "L"
'						ShippingRequired = "No"
'				End Select
'		  	Next
'		  	'if no books then assume nothing needs to be shipped
'		  	If NumBooks = 0 Then ShippingRequired = "No"
'  	End Select
  	
'	If ShippingRequired = "Yes" Then





        'new shipping method
        'shipping cost = [Price of 1 book] + [(n-1) * (Price of 3 books - Price of 1 book)/2]
        'table: ShippingCalculator
        'fields: Country | Continent | 1Book | 3Books | TransitTime | DateUpdated
            Dim bk
            Dim totalWeight
            Dim bookKeys
            pkgWeight = 0
            
            ' Calculate the total weight of all the books being shipped
            bookKeys = BookList.Keys
            For bk = 0 To BookList.Count - 1

                Dim bookWeight
                bookWeight = GetBookWeight(bookKeys(bk), "lb")
                
                If bookWeight > 0 Then
                    pkgWeight = pkgWeight + BookList.Item(bookKeys(bk)) * bookWeight
                Else
                    pkgWeight = -1
                    Exit For
                End if
WriteLogEntry( "tw: " & pkgWeight & ", key: " & bookKeys(bk) & " item: " & BookList.Item(bookKeys(bk)) & "<br/>")  		
            Next
          
            If pkgWeight > 0 Then
                ' Get the shipping estimate from Purolator
                ShippingCost = PdsGetLowestExpressPrice( _
                                                "N1H4E9", _
                                                ClientCity, _
                                                ClientProvince, _
                                                ConvertCountryToCountryCode(ClientCountry), _
                                                ClientPostalCode, _
                                                "CustomerPackaging", _
                                                Round(pkgWeight), _
                                                "lb")
WriteLogEntry("shippingCost: " & ShippingCost)
                Session("ShippingEstimateMethod") = "Purolator Online, total weight " & Round(pkgWeight) & " lbs: $" & ShippingCost & " + $10 handling"
                'add overhead
		's if something goes wrong in the method above, we might get -1 as the cost value so check below to ensure everything goes right
		If ShippingCost >0 Then
		    ShippingCost = ShippingCost + 10
		Else 
		    ShippingCost = 0
		End If
            Else
                ShippingCost = 0
            End If
    
                
            ' If the web service fails, use the old calculating method
            ' Also, always use this method for shipments to the USA.
            If ShippingCost <= 0 Or CurrencyStr = "USA" Then
                Session("ShippingEstimateMethod") = "CHI Generated"
                Select Case CurrencyStr
                    Case "Canadian"
                        '=============================================================================================
                        '     Books shipped by Purolator Ground from CHI HQ                            - CAD$24 + 2i
                        '=============================================================================================
                        If NumBooks > 0 Then
                                ShippingCost = 24 + (2 * NumBooks)
                                If ShipWorkbook Then ShippingCost = ShippingCost + 21
                        End If
                    Case "USA"
                        '=============================================================================================
                        '     Books shipped by UPS Ground from US Address                               - US$15 + 3i
                        '=============================================================================================
                        If NumBooks > 0 Then
                                ShippingCost = 15 + (3 * NumBooks)
                                If ShipWorkbook Then ShippingCost = ShippingCost + 28
                        End If
                    Case "International"
                        '=============================================================================================
                        '     Books shipped by Xpress Post* from CHI HQ                                 - US$26 + 15i
                        '     * use International Air Parcel where Xpress Post unavailable
                        '=============================================================================================
                        If NumBooks > 0 Then
                                ShippingCost = 26 + (15 * NumBooks)
                                If ShipWorkbook Then ShippingCost = ShippingCost + 28
                        End If
                End Select
            End If
                
                'new method where we tie into purolator's web services
		
WriteLogEntry("ShippingEstimateMethod: " & Session("ShippingEstimateMethod"))
WriteLogEntry("ShippingCost total: " & ShippingCost)       
	'Else
	'	ShippingCost = 0
	'End If
	
	'Session("ShippingCalculated") = "Yes"
	'
	''if backend user then mark changes
	'If Session("PasswordOK") = "Yes" and Session("CurrentTransaction") <> 0 then Session("Edited") = True
	'
	'Response.Redirect Session("NextPage")
	'
	If ShippingCost <=0 Then
	    CalculateShipping = 0
	Else
	    CalculateShipping = ShippingCost
	End If
	
End Function

%>





 