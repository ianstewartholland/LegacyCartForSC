<%
    Public Function RetrievePaperInfo(id)
    	Dim objConn, strSQL, objRS, objCommand, objDbParam
	Dim strDBConnect

	strDBConnect = "Driver={Microsoft Access Driver (*.mdb)};DBQ=C:\WWWUsers\LocalUser\chi\computationalhydraulics_com\databases\biblio.mdb"

	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionTimeout = 15
	objConn.CommandTimeout = 30
	objConn.Open = strDBConnect
	
	strSQL = "SELECT ID,AUTHOR,TITLE,SOURCE,PLACE,BOOK,PAGE,DATE,PUBDATE,ABSTRACT FROM C152"
	strSQL = strSQL & " WHERE ID= ?;"
	
	set objCommand = Server.CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConn
	objCommand.CommandText = strSQL
	
	set objDbParam = objCommand.CreateParameter("@paperid", 200, 1, 10)
	objCommand.Parameters.Append objDbParam
	objCommand.Parameters("@paperid") = id
	set objDbParam = Nothing

	set objRS = objCommand.Execute()
	
	if objRS.BOF or objRS.EOF then
		RetrievePaperInfo = Empty
	else
		RetrievePaperInfo = objRS.GetRows()
	end if
	
	set objCommand = Nothing
	objConn.Close
        set objConn = Nothing
    End Function

' Expects the following session variables to have been set:
'   CurrentItem - 
'   LastPage - last page user was viewing
 

	dim Item
	dim NumItems 
	dim t
	dim n
        dim ItemInfo
        
        Item = Request.QueryString("item")
        If Len(Item) > 1 Then
            Session("CurrentItem") = Item
        End If
        
        'check for expired page
	If Session("CurrentItem") = "" Then Response.Redirect "cart.asp"
        
        ' retrieve the item information from the publications database.
        ItemInfo = RetrievePaperInfo(Session("CurrentItem"))
        If isEmpty(ItemInfo) Then Response.Redirect "cart.asp"

  	'store this page
  	Session("NextPage") = "addbooktocartcommit.asp"
	
	'check to ensure the Catalog is loaded
	'if Session("CatalogLoaded") <> "Yes" then Response.Redirect("Catalog.asp")
        
	'set the return page
'        Select Case Left(Session("CurrentItem"),1)
'		Case "R"
'			Session("LastPage") = "http://www.computationalhydraulics.com/Publications/Books/"
'		Case "S"
'			Session("LastPage") = "http://www.computationalhydraulics.com/Software/PCSWMM.NET/"
'		Case "W"
'			Session("LastPage") = "http://www.computationalhydraulics.com/Training/Workshops/"
'		Case "C"
'			Session("LastPage") = "http://www.computationalhydraulics.com/Training/Conferences/"
'		Case Else
'			Session("LastPage") = "http://www.computationalhydraulics.com"
'	End Select
        
        
	NumItems = 0
	
	'initialize the total number of items
	'If Session("NumCartItems") > 0 Then
	'Else
	'	Session("TotalNumItems") = 0
	'End If
	
	' only need one copy of the paper
	NumItems = 1 
	
	'check that Item exists
        Item = Session("CurrentItem")
	If len(Item) > 1 Then
	
		'check if item is already in cart only allow a quantity of 1
		For t = 1 to Session("NumCartItems")
			If Session("CartItem(" & t & ")") = Item then
				Session("CartItemQuantity(" & t & ")") = CInt(NumItems)
				Exit For
			End If
		Next

		'if not in the cart then add item
		If t > Session("NumCartItems") Then
			Session("CartItem(" & t & ")") = UCase(CStr(Item))
			Session("CartItemQuantity(" & t & ")") = NumItems
			Session("CartItemDescription(" & t & ")") = ItemInfo(2,0)
			Session("CartItemCost(" & t & ")") = 14.95 ' hard wired for now.
			Session("NumCartItems") = t
		End If

		'erase current item
		Session("CurrentItem") = ""

		'Add to the count of the total number of items
		'Session("TotalNumItems") = Session("TotalNumItems") + NumItems
		
	End If
	
	'set flags for recalculating
	'Session("SubTotalCalculated") = ""
	'Session("ShippingCalculated") = ""
	'Session("ShippingRequired") = ""
	'Session("TaxCalculated") = ""
	'Session("TotalCost") = 0

	Response.Redirect "cart.asp"
			
%>








 