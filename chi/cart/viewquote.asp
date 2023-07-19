
<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<%
    ' do not check regular password but verify order info before display
	If Len(Request.QueryString("Transaction")) > 1 And Len(Request.QueryString("LastName")) > 1 Then
            CurrentTransaction = Request.QueryString("Transaction")
            Email = Request.QueryString("LastName")
	Else
            Response.End
	End If

    	' set connection object here
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
	conn.Open transcationPath

	'connect to database and load the information for this page
	Set objRec = Server.CreateObject ("ADODB.Recordset")	
	'objRec.Open "Transaction", transcationPath, 3, 3, 2
    objRec.Open "[Transaction]", conn
	objRec.Filter = "TransactionNumber = '" & CurrentTransaction & "'"

    If objRec.EOF OR objRec.BOF Then
        objRec.Close
	    Set objRec = Nothing	
        Response.End
    End If

	objRec.MoveFirst

	'load client info ===========================

	'Session("ClientAddress") = "Yes"
	ClientPrefix = ObjRec("ClientPrefix")
	ClientFirstName = ObjRec("ClientFirstName")
	ClientLastName = ObjRec("ClientLastName")
	ClientName =  Trim(ClientPrefix & " " & Trim(ClientFirstName & " " & ClientLastName))
	ClientCompany = ObjRec("ClientCompany")
	ClientStreetAddress = ObjRec("ClientStreetAddress")
	ClientStreetAddress2 = ObjRec("ClientStreetAddress2")
	ClientCity = ObjRec("ClientCity")
	ClientProvince = ObjRec("ClientProvince")
	ClientCountry = ObjRec("ClientCountry")
	ClientPostalCode = ObjRec("ClientPostalCode")
	ClientEmail = ObjRec("ClientEmail")
	ClientTelephone = ObjRec("ClientTelephone")
	ClientFax = ObjRec("ClientFax")
	ClientWebSite = ObjRec("ClientWebSite")
	ClientType = ObjRec("ClientType")
	'PreviousLicensee = ObjRec("PreviousLicensee")

    If LCase(Trim(ClientLastName)) <> LCase(Trim(Email)) Then
        objRec.Close
	    Set objRec = Nothing	
        Response.End
    End If

	'load shipping info ===========================

	'Session("ShipAddress") = "Yes"
	ShipPrefix = ObjRec("ShipPrefix")
	ShipFirstName = ObjRec("ShipFirstName")
	ShipLastName = ObjRec("ShipLastName")
	ShipName =  Trim(ShipPrefix & " " & Trim(ShipFirstName & " " & ShipLastName))
	ShipCompany = ObjRec("ShipCompany")
	ShipStreetAddress = ObjRec("ShipStreetAddress")
	ShipStreetAddress2 = ObjRec("ShipStreetAddress2")
	ShipCity = ObjRec("ShipCity")
	ShipProvince = ObjRec("ShipProvince")
	ShipCountry = ObjRec("ShipCountry")
	ShipPostalCode = ObjRec("ShipPostalCode")
	ShipEmail = ObjRec("ShipEmail")
	ShipTelephone = ObjRec("ShipTelephone")
	ShipFax = ObjRec("ShipFax")
	
	'load billing info ===========================

	'Session("BillAddress") = "Yes"
	BillPrefix = ObjRec("BillPrefix")
	BillFirstName = ObjRec("BillFirstName")
	BillLastName = ObjRec("BillLastName")
	BillName =  Trim(BillPrefix & " " & Trim(BillFirstName & " " & BillLastName))
	BillCompany = ObjRec("BillCompany")
	BillStreetAddress = ObjRec("BillStreetAddress")
	BillStreetAddress2 = ObjRec("BillStreetAddress2")
	BillCity = ObjRec("BillCity")
	BillProvince = ObjRec("BillProvince")
	BillCountry = ObjRec("BillCountry")
	BillPostalCode = ObjRec("BillPostalCode")
	BillEmail = ObjRec("BillEmail")
	BillTelephone = ObjRec("BillTelephone")
	BillFax = ObjRec("BillFax")

	'load currencies ======================================
	
	CurrencyStr = ObjRec("Currency")
	CurrencyAbrev = ObjRec("CurrencyAbrev")
	If IsNull(ObjRec("CurrencyExchange")) Then
	    CurrencyExchange = -1 ' put -1 in case the currency has not been set, people would be able to tell
	Else
	    CurrencyExchange = ObjRec("CurrencyExchange")
	End If
	
'	CurrencyExchange = 0 ' assign default value
'	CurrencyExchange = ObjRec("CurrencyExchange")
'        If IsNull(CurrencyExchange) Then CurrencyExchange = 0
	
	'ShippingCalculated = "Yes"
	ShippingCost = ObjRec("ShippingCost")
        If IsNull(ShippingCost) Then ShippingCost = 0

	'TaxCalculated = "Yes"
        HSTCost = ObjRec("HSTCost")
        If IsNull(HSTCost) Then HSTCost = 0
	GSTCost = ObjRec("GSTCost")
        If IsNull(GSTCost) Then GSTCost = 0
	PSTCost = ObjRec("PSTCost")
        If IsNull(PSTCost) Then PSTCost = 0

	'Session("SubTotalCalculated") = "Yes"
	SubTotalCost = ObjRec("SubTotalCost")
        If IsNull(SubTotalCost) Then SubTotalCost = 0
	TotalCost = ObjRec("TotalCost")
        If IsNull(TotalCost) Then TotalCost = 0
	OriginalTotalCost = ObjRec("OriginalTotalCost")
        If IsNull(OriginalTotalCost) Then OriginalTotalCost = 0

	'load dates ===========================================

	OrderDate = ObjRec("OrderDate")

	PurchaseOrder = objRec("PurchaseOrder")

	'load items ===========================================
	Items = ObjRec("Items")
	Dim Item
	Dim ItemDescription
	Dim ItemCost
	Dim ItemQuantity
	n=1
	
	'define array to store the item information
	Dim CartItem()
	Dim CartItemDescription()
	Dim CartItemCost()
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
			
			redim preserve CartItemDescription(t)
			CartItemDescription(t) = ItemDescription
			
			redim preserve CartItemCost(t)
			CartItemCost(t) = ItemCost
			
			redim preserve CartItemQuantity(t)
			CartItemQuantity(t) = ItemQuantity

			If n > len(Items) Then Exit For
		Next
	End If
	
	objRec.Close
	Set objRec = Nothing	

%>
<html>

<head>
<title>CHI Quote #<%= CurrentTransaction %></title>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache"> <% 'force page refresh %>
<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<META HTTP-EQUIV="EXPIRES" CONTENT="0">
<!meta description="Consulting engineering firm specializing in urban stormwater quantity and quality management modeling esp SWMM">
<!meta keywords="SWMM USEPA SWMM4.4 PCSWMM stormwater management model modeling modelling BMP's catchbasin civil engineering channel computer conduit discharge hydraulics hydrology management modeling modelling NPDES non-pointsource pipe pollutants precipitation programming quality rain rainfall runoff sediment simulation storm stormwater streamflow transport urban water quality water resources engineering">
</head>

<body text="#000000" bgcolor="#FFFFFF" link="#0000EE" vlink="#551A8B" alink="#FF0000">

<p>
<table border="0" width="665" cellpadding="0" cellspacing="0">
  <tr>
    <td width="365" valign="top" colspan="2">
        <p><font face="Arial" size="5"><b>Computational Hydraulics Int.</b></font></p>
    </td>
    <td width="300" align="right" valign="top">
    <font face="Arial" size="5"><b>Quote <% Response.Write CurrentTransaction %></b></font>
    </td>
  </tr>
  <tr>
    <td width="50" valign="top">
    </td>
    <td width="270" valign="top">
        <font face="Arial" size="2">
        147 Wyndham Street, Suite 202<br>
	Guelph, Ontario, Canada, N1H 4E9<br>
        Tel: (519) 767-0197 Fax: (519) 489-0695<br>
        Email: <a href="mailto:info@chiwater.com">info@chiwater.com</a><br>
        Web: www.chiwater.com</font>
    </td>
    <td width="300" align="right" valign="top">
    <font face="Arial" size="2"><% Response.Write FormatDateTime(OrderDate,1) %></font>
    </td>
  </tr>
</table>

<font face="Arial">

<br>

</font>

<table border="0" width="620">
  <tr>
    <td width="620" colspan="3" valign="top">
    <font face="Arial" size="2"><b>
    Client Information</b></font>
    </td>
  </tr>
  <tr>
    <td width="50" valign="top">
    </td>
    <td width="270" valign="top"><font size="2" face="Arial">
    	<%
    	Response.Write ClientName & "<br>"
    	Response.Write ClientCompany & "<br>"
    	Response.Write ClientStreetAddress & "<br>"
     	If ClientStreetAddress2 <> "" Then Response.Write ClientStreetAddress2 & "<br>"
    	Response.Write ClientCity & " " & ClientProvince & " " & ClientPostalCode & "<br>"
    	Response.Write ClientCountry & "<br><br>"
    	%></font>
    </td>
    <td width="300" valign="top"><font size="2" face="Arial">
    	<%
    	Response.Write "Email: " & ClientEmail & "</a><br>"
    	Response.Write "Tel: " & ClientTelephone & "<br>"
    	Response.Write "Fax: " & ClientFax & "<br><br>"
    	%></font>
    </td>
  </tr>
  <tr>
    <td width="620" colspan="3" valign="top">
    <font face="Arial" size="2"><b>
    Shipping Address</b></font>
    </td>
  </tr>
  <tr>
    <td width="50" valign="top">
    </td>
    <% 
    'check if shipping address is the same as the client address
    ShippingSame = "Yes"
    If ShipName <> ClientName Then ShippingSame = "No"
    If ShipCompany <> ClientCompany Then ShippingSame = "No"
    If ShipStreetAddress <> ClientStreetAddress Then ShippingSame = "No"
    If ShipCity <> ClientCity Then ShippingSame = "No"
    If ShipProvince <> ClientProvince Then ShippingSame = "No"
    If ShipPostalCode <> ClientPostalCode Then ShippingSame = "No"
    If ShipCountry <> ClientCountry Then ShippingSame = "No"
    If ShipEmail <> ClientEmail Then ShippingSame = "No"
    If ShipTelephone <> ClientTelephone Then ShippingSame = "No"
    If ShipFax <> ClientFax Then ShippingSame = "No"
    If ShippingSame = "Yes" Then
    %>
    <td width="570" colspan="2" valign="top"><font size="2" face="Arial">
    Same as client address above...<br><br>
    </font>
    </td>
    <%
    Else
    %>
    <td width="270" valign="top"><font size="2" face="Arial">
    	<%
    	Response.Write ShipName & "<br>"
    	Response.Write ShipCompany & "<br>"
    	Response.Write ShipStreetAddress & "<br>"
    	If ShipStreetAddress2 <> "" Then Response.Write ShipStreetAddress2 & "<br>"
    	Response.Write ShipCity & " " & ShipProvince & " " & ShipPostalCode & "<br>"
    	Response.Write ShipCountry & "<br><br>"
    	%></font>
    </td>
    <td width="300" valign="top"><font size="2" face="Arial">
    	<%
    	Response.Write "Email: " & ShipEmail & "<br>"
    	Response.Write "Tel: " & ShipTelephone & "<br>"
    	Response.Write "Fax: " & ShipFax & "<br><br>"
    	%></font>
    </td>
    <%
    End If
    %>
  </tr>
  <tr>
    <td width="620" colspan="3" valign="top">
    <font face="Arial" size="2"><b>
    Billing Address</b></font>
    </td>
  </tr>
  <tr>
    <td width="50" valign="top">
    </td>
    <% 
    'check if billing address is the same as the client address
    BillingSame = "Yes"
    If BillName <> ClientName Then BillingSame = "No"
    If BillCompany <> ClientCompany Then BillingSame = "No"
    If BillStreetAddress <> ClientStreetAddress Then BillingSame = "No"
    If BillStreetAddress2 <> ClientStreetAddress2 Then BillingSame = "No"
    If BillCity <> ClientCity Then BillingSame = "No"
    If BillProvince <> ClientProvince Then BillingSame = "No"
    If BillPostalCode <> ClientPostalCode Then BillingSame = "No"
    If BillCountry <> ClientCountry Then BillingSame = "No"
    If BillEmail <> ClientEmail Then BillingSame = "No"
    If BillTelephone <> ClientTelephone Then BillingSame = "No"
    If BillFax <> ClientFax Then BillingSame = "No"
    If BillingSame = "Yes" Then
    %>
    <td width="570" colspan="2" valign="top"><font size="2" face="Arial">
    Same as client address above...<br><br>
    </font>
    </td>
    <%
    Else
    %>
    <td width="270" valign="top"><font size="2" face="Arial">
    	<%
    	Response.Write BillName & "<br>"
     	If BillCompany <> "" Then Response.Write BillCompany & "<br>"
    	Response.Write BillStreetAddress & "<br>"
    	If BillStreetAddress2 <> "" Then Response.Write BillStreetAddress2 & "<br>"
    	Response.Write BillCity & " " & BillProvince & " " & BillPostalCode & "<br>"
    	Response.Write BillCountry & "<br><br>"
    	%></font>
    </td>
    <td width="300" valign="top"><font size="2" face="Arial">
    	<%
    	Response.Write "Email: " & BillEmail & "<br>"
    	Response.Write "Tel: " & BillTelephone & "<br>"
    	Response.Write "Fax: " & BillFax & "<br><br>"
    	%></font>
    </td>
    <%
    End If
    %>
  </tr>
    <%
    If PurchaseOrder <> "" Then
	    %>
		<tr>
			<td width="620" colspan="3" valign="top"><font face="Arial" size="2"><b>Purchase Order #</b></font></td>
		</tr>
		<tr>
			<td width="50" valign="top"></td>
			<td width="270" valign="top"><font size="2" face="Arial"><% Response.Write PurchaseOrder & "<br><br>" %></font></td>
			<td width="300" valign="top"></td>
		</tr>
	    <%
    End If
    %>
</table>
<table border="0" width=620>
  <tr>
    <td width="45" valign="top"><b><font size="2" face="Arial">Item</font></b></td>
    <td width="365" valign="top"><font size="2" face="Arial"><b>Description</b></font></td>
    <td width="70" align="right" valign="top">
      <p align="right"><b><font size="2" face="Arial">Unit Price</font></b></p>
    </td>
    <td width="70" align="right" valign="top"><b><font size="2" face="Arial">Quantity</font></b></td>
    <td width="70" align="right" valign="top"><b><font size="2" face="Arial">Amount</font></b></td>
  </tr>
  <%
	isConference = False
	
  	For t = 1 to NumCartItems
  		TotalItemCost = CartItemCost(t) * CartItemQuantity(t)
  %>
  <tr>
    <td width="45" valign="top">
      <p align="left"><font size="2" face="Arial">
      <%
	response.write CartItem(t)
      
        ' for conference year 2012, we extend the quote exp date till the conference date - change below for next year or ignore this
	' date this was done Nov 21, 2011 - so we might have 4 quotes sent out already before this date
	' if you do change below, remember to change date in the very bottom
	' I am not lazy to look the item/date up in the database. But classic asp has an issue with the database connection after a server upgrade...yeah, I could..blah - just to save us from more trouble
	' the other thing is the payqote page should also be changed accordingly like below
	If CartItem(t) = "C021" Or CartItem(t) = "C321" Or CartItem(t) = "C421" Then isConference = True
      
      %></font></td>
    <td width="365" valign="top">
      <p align="left"><font size="2" face="Arial">
      <% response.write CartItemDescription(t) %></font></td>
    <td width="70" align="right" valign="top">
      <p align="right"><font size="2" face="Arial">
      <% response.write FormatCurrency(CartItemCost(t),2) %></font></td>
    <td width="70" align="right" valign="top">
      <p align="right"><font size="2" face="Arial">
      <% response.write CartItemQuantity(t) %></font></td>
    <td width="70" align="right" valign="top">
      <p align="right"><font size="2" face="Arial">
      	<% response.write FormatCurrency(TotalItemCost ,2) %>
	</font></td>
  </tr>
<%
  	Next
%>
  <tr>
    <td width="45" valign="top">
	</td>
    <td width="365" valign="top">
      </td>
    <td width="70" align="right" valign="top">
    </td>
    <td width="70" align="right" valign="top">
    <p align="right"><b><font size="2" face="Arial">SubTotal:</font></b></td>
    <td width="70" align="right" valign="top">
      <p align="right"><font size="2" face="Arial">
      <% response.write FormatCurrency(SubTotalCost,2) %></font></td>
  </tr>
  <%
  		If GSTCost > 0 Then
  %>
  <tr>
    <td width="45" valign="top">
	</td>
    <td width="365" valign="top">
      </td>
    <td width="140" align="right" valign="top" colspan="2">
    <p align="right"><b><font size="2" face="Arial">GST (#119880979):</font></b></td>
    <td width="70" align="right" valign="top">
      <p align="right"><font size="2" face="Arial">
      <% response.write FormatCurrency(GSTCost,2) %></font></td>
  </tr>
  <%
  		End If
		If PSTCost > 0 Then
  %>
  <tr>
    <td width="45" valign="top">
	</td>
    <td width="365" valign="top">
      </td>
    <td width="70" align="right" valign="top">
    </td>
    <td width="70" align="right" valign="top">
    <p align="right"><b><font size="2" face="Arial">
      PST:</font></b></td>
    <td width="70" align="right" valign="top">
      <p align="right"><font size="2" face="Arial">
      <% response.write FormatCurrency(PSTCost,2) %></font></td>
  </tr>
  <%
  		End If
		If HSTCost > 0 Then
  %>
  <tr>
    <td width="60" valign="top">
	</td>
    <td width="365" valign="top">
      </td>
    <td width="70" align="right" valign="top">
    </td>
    <td width="70" align="right" valign="top">
    <p align="right"><b><font size="2" face="Arial">
      HST:</font></b></td>
    <td width="70" align="right" valign="top">
      <p align="right"><font size="2" face="Arial">
      <% response.write FormatCurrency(HSTCost,2) %></font></td>
  </tr>
  <%
  		End If
  %>
  <tr>
    <td width="45" valign="top">
	</td>
    <td width="505" valign="top" colspan="3">
    <p align="right"><b><font size="2" face="Arial">
      Total in <% Response.Write CurrencyAbrev %> dollars:</font></b>
      </td>
    <td width="70" align="right" valign="top">
      <p align="right"><font size="2" face="Arial"><b>
      <% response.write FormatCurrency(TotalCost,2) %></b></font></td>
  </tr>
</table>
<p><font size="2" face="Arial">To confirm this order and enter your payment information, please go to:<br><a href=http://www.chiwater.com/pay.asp<% Response.Write "?Item=" & CurrentTransaction & "&Name=" & ClientLastName %>>http://www.chiwater.com/pay.asp<% Response.Write "?Item=" & CurrentTransaction & "&Name=" & ClientLastName  %></a></font></p>

<table border="0" width="665">
    <tr>
	<td>
<p align="center"><br /><font size="2" face="Arial">
<%
' get the expire quote date
newDate = DateAdd("m", 2, OrderDate)

'If Not isConference Then
%>
    This quote is valid till <%= MonthName(Month(newDate)) & " " & Day(newDate) & ", " & Year(newDate) %>.<br>
<% 'Else 
  '  This quote is valid till February 22, 2012.<br>
' End If %>
<br>
If you have any questions concerning this quote, call: (519) 767-0197<br>
<br>
<b>THANK YOU FOR YOUR BUSINESS!</b></font></p>
	</td>
    </tr>

</body>
</html>




 