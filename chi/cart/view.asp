<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<%
'
' important, watch out for sql/script injection
' watch out for single quotation mark, space; future steps : and please valid user's input
'
Function IsValid(input)
    pass = False
    
    ' for transaction numbers
    If IsNumeric(input) And Len(input) < 6 Then
	If InStr(input, " ") = 0 And InStr(input, "'") = 0 Then
	    pass = True
	Else
	    pass = False
	End If 
    End If
    
'    If InStr(input, "'") > 0 Then
'	' find the ' and replace it with ''
'	
'	
'	
'
'    End If
'    
'    Dim blackList(2)
'    blackList(1) = "script"
'    blackList(2) = "ftp"  
    
    IsValid = pass
    
End Function

Session("QuoteQuery") = "PCSWMM"

	If IsValid(Trim(Request.QueryString("Item"))) Then 
	    CurrentTransaction = Trim(Request.QueryString("Item"))
	Else
	    Response.Redirect "http://www.chiwater.com/view.asp?Error=Error"
	End If
	
	Set objRec = Server.CreateObject ("ADODB.Recordset")	
	objRec.Open "Transaction", transcationPath, 3, 3, 2
	objRec.Filter = "TransactionNumber = '" & CurrentTransaction & "'"
	
	If objRec.EOF And objRec.BOF Then
	    Response.Redirect "http://www.chiwater.com/view.asp?Error=NotFound"
	Else
	
	    objRec.MoveFirst
	    
	    ' check valid client 
	    If LCase(ObjRec("ClientLastName")) <> LCase(Trim(Request.QueryString("Name"))) Then Response.Redirect "http://www.chiwater.com/view.asp?Error=Name"
	    
	    ' data check
	    If Not (DateDiff("d", DateAdd("m", 6, ObjRec("PaidDate")), Date) < 1) Then Response.Redirect "http://www.chiwater.com/view.asp?Error=Expired"
	    
	    If objRec("Cancelled") = "Yes" Then Response.Redirect "http://www.chiwater.com/view.asp?Error=Error"
	    
	    'load client info ===========================
    
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
    
	    'load shipping info ===========================

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
	    
	    'load payment info ===================================
	    
	    CardType = ObjRec("CardType")
	    Select Case CardType
		    Case "VISA", "Mastercard", "American Express", "Wire transfer"
		    Case Else
			    CardType = ""
	    End Select
	    CardNumber = ObjRec("CardNumber")
	    CardName = ObjRec("CardName")
	    ChequeNumber = ObjRec("ChequeNumber")
	    'load currencies ======================================
	    
	    CurrencyAbrev = ObjRec("CurrencyAbrev")

	    ShippingCost = ObjRec("ShippingCost")
	    ShippingRequired = ObjRec("ShippingRequired")
	    If IsNull(ShippingCost) Then ShippingCost = 0
    
	    HSTCost = ObjRec("HSTCost")
	    If IsNull(HSTCost) Then HSTCost = 0
	    GSTCost = ObjRec("GSTCost")
	    If IsNull(GSTCost) Then GSTCost = 0
	    PSTCost = ObjRec("PSTCost")
	    If IsNull(PSTCost) Then PSTCost = 0
    
	    SubTotalCost = ObjRec("SubTotalCost")
	    If IsNull(SubTotalCost) Then SubTotalCost = 0
	    TotalCost = ObjRec("TotalCost")
	    If IsNull(TotalCost) Then TotalCost = 0
    
	    'load dates ===========================================
    
	    OrderDate = ObjRec("OrderDate")
	    PaidDate = ObjRec("PaidDate")
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
    
			    ' set quote control here
			    If LCase(Left(Item, 1)) = "w" Then
				Session("QuoteQuery") = "Workshops"
			    End If 
    
			    If n > len(Items) Then Exit For
		    Next
	    End If
	End If
	objRec.Close
	Set objRec = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- #BeginTemplate "/Templates/innerpage2.dwt" --><!-- DW6 -->

<!-- #BeginEditable "variables" -->
<% 
Dim mainpage, subpage
mainpage = ""
subpage = "View Transaction"
%>
<!-- #EndEditable -->

<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<link href="styles/style2.css" rel="STYLESHEET" type="text/css" media="screen" />
	<link href="styles/print.css" rel="STYLESHEET" type="text/css" media="print" />
	<link rel="shortcut icon" href="images/chilogo.ico" />
	<!-- #BeginEditable "title" -->
<title>View Transaction - CHI</title>
<!-- #EndEditable -->
<script type="text/javascript" src="styles/cart.js"></script>
</head>
<body onload='loadButton()'>
<div class="main">
		<div class="logo_home">			
			<div class="header">
				<a style="width:150px;height:100px;float:left;text-decoration:none;" href="http://www.chiwater.com">&nbsp;</a>
				<div class="slogan">
					<!-- #BeginEditable "slogan" -->
					    Dedicated to promoting healthier ecosystems and communities<br />
					    by providing the software, resources and guidance to simplify<br />
					    advanced water management and design.
					<!-- #EndEditable -->
				</div>
<div class="print_search">
	<form action="http://www.chiwater.com/search_results.asp" id="searchform">
<%
If Session("NumCartItems") <> "" Then
    If Session("NumCartItems") > 1 Then
        Response.Write "<a style='color: #FFFFFF; link: #FFFFFF; visited: #FFFFFF; font-size:11px; font-weight:bold;' class='nav' href = 'https://secure.chiwater.com/chi/cart/cart.asp'>Cart (" & Session("NumCartItems") & " items)</a>&nbsp;&nbsp;"
    ElseIf Session("NumCartItems") = 1 Then
        Response.Write "<a style='color: #FFFFFF; link: #FFFFFF; visited: #FFFFFF; font-size:11px; font-weight:bold;' class='nav' href = 'https://secure.chiwater.com/chi/cart/cart.asp'>" &  "Cart (1 item)</a>&nbsp;&nbsp;"
    End If
End If
%>
		<input type="text" name="q" value="Search..." class="searchbox" onfocus="javascript:this.value='';" onblur="javascript:this.value='Search...';this.style.color = '#808080';"/>
		<input type="image" src="images/search_arrow.png" style="vertical-align:top;"/>
	</form>
</div>
			</div>			
			
			<div style="clear:both;"></div>
		</div>
	<!-- #include file = "styles/topmenu2.asp" -->
		<div class="banner"></div>
		<div class="innerbg">
			<ul>
				
				<li class="main_content_large">
				
				<!-- #BeginEditable "Body" -->
					<ul>
						<li class="maincontenttop_large"></li>
						<li class="maincontentarea_large">	
<form>
    <div id="printContent">
<table width="96%">
  <tr>
    <td valign="top" colspan="2">
        <span style="padding-left:0px; padding-right:0px;margin-bottom:0px;color:#5E3F22;font-size:20px;font-weight:bold;">Computational Hydraulics Int.</span>
    </td>
    <td valign="top" align = "right" colspan="3">
    <span style="padding-right:0px;margin-bottom:0px; float:right;color:#5E3F22;font-size:20px;font-weight:bold;">
    <%
    word = ""
    If IsDate(PaidDate) Then 
	word = "Statement "
    Else
	word = "Invoice "
    End If
    
    Response.Write word
    %>
     <% Response.Write CurrentTransaction %></span>
    </td>
  </tr>
  <tr>
    <td width = '8%' valign="top">
    </td>
    <td valign="top" >
        147 Wyndham Street, Suite 202<br>
	Guelph, Ontario, Canada, N1H 4E9<br>
        Tel: (519) 767-0197 Fax: (519) 489-0695<br>
        Email: <a href="mailto:info@chiwater.com">info@chiwater.com</a><br>
        Web: www.chiwater.com
	<br><br>
    </td>
    <td align="right" valign="top" colspan="3">
    <%
    If IsDate(PaidDate) Then 
	Response.Write FormatDateTime(Date,1)
    Else
	Response.Write FormatDateTime(OrderDate,1)
    End If
    
    
     %>&nbsp;
    </td>
  </tr>

  <tr>
    <td colspan="3" valign="top">
    <b>
    Client Information</b>
    </td>
  </tr>
  <tr>
    <td width="8%" valign="top">
    </td>
    <td width="270" valign="top">
    	<%
    	Response.Write ClientName & "<br>"
    	Response.Write ClientCompany & "<br>"
    	Response.Write ClientStreetAddress & "<br>"
     	If ClientStreetAddress2 <> "" Then Response.Write ClientStreetAddress2 & "<br>"
    	Response.Write ClientCity & " " & ClientProvince & " " & ClientPostalCode & "<br>"
    	Response.Write ClientCountry & "<br><br>"
    	%>
    </td>
    <td width="300" valign="top" >
    	<%
    	Response.Write "Email: " & ClientEmail & "<br>"
    	Response.Write "Tel: " & ClientTelephone & "<br>"
    	Response.Write "Fax: " & ClientFax 
    	%>
    </td>
  </tr>
<% If ShippingRequired = "Yes" Then %>
  <tr>
    <td colspan="3" valign="top">
    <b>
    Shipping Address</b>
    </td>
  </tr>
  <tr>
    <td valign="top">
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
    <td width="570" colspan="2" valign="top">
    Same as client address above...
    <br><br>
    </td>
    <%
    Else
    %>
    <td width="270" valign="top">
    	<%
    	Response.Write ShipName & "<br>"
    	Response.Write ShipCompany & "<br>"
    	Response.Write ShipStreetAddress & "<br>"
    	If ShipStreetAddress2 <> "" Then Response.Write ShipStreetAddress2 & "<br>"
    	Response.Write ShipCity & " " & ShipProvince & " " & ShipPostalCode & "<br>"
    	Response.Write ShipCountry & "<br><br>"
    	%>
    </td>
    <td width="300" valign="top" >
    	<%
    	Response.Write "Email: " & ShipEmail & "<br>"
    	Response.Write "Tel: " & ShipTelephone & "<br>"
    	Response.Write "Fax: " & ShipFax 
    	%>
    </td>
    <%
    End If
    %>
  </tr>
<% End If %>
  <tr>
    <td width="96%" colspan="3" valign="top">
    <b>
    Billing Address</b>
    </td>
  </tr>
  <tr>
    <td valign="top">
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
    <td width="570" colspan="2" valign="top">
    Same as client address above...
    <br><br>
    </td>
    <%
    Else
    %>
    <td width="270" valign="top">
    	<%
    	Response.Write BillName & "<br>"
     	If BillCompany <> "" Then Response.Write BillCompany & "<br>"
    	Response.Write BillStreetAddress & "<br>"
    	If BillStreetAddress2 <> "" Then Response.Write BillStreetAddress2 & "<br>"
    	Response.Write BillCity & " " & BillProvince & " " & BillPostalCode & "<br>"
    	Response.Write BillCountry & "<br><br>"
    	%>
    </td>
    <td width="300" valign="top" >
    	<%
    	Response.Write "Email: " & BillEmail & "<br>"
    	Response.Write "Tel: " & BillTelephone & "<br>"
    	Response.Write "Fax: " & BillFax 
    	%>
    </td>
    <%
    End If
    %>
  </tr>
    <%
    If PurchaseOrder <> "" Then
	    %>
		<tr>
			<td width="620" colspan="3" valign="top"><b>Purchase Order #</b></td>
		</tr>
		<tr>
			<td valign="top"></td>
			<td width="270" valign="top"><% Response.Write PurchaseOrder %><br><br></td>
			<td width="300" valign="top"></td>
		</tr>

      <% End If %>
  
    <%
    If IsDate(PaidDate) Then
    %>
      <tr>
	<td width="620" colspan="3" valign="top">
	<b>Payment Method</b>
	</td>
      </tr>      
	<%If CardNumber <> "" Then 
	%>

      <tr>
	<td valign="top">
	</td>
	<td width="270" valign="top" colspan="2">
	    <%
		Response.Write "Card type: " & CardType & "<br>"
		Response.Write "Cardholder name: " & CardName & "<br><br>"
	    %>
	</td>

      </tr>
      <% End If
	If ChequeNumber <> "" Then
	    If CardType = "Wire transfer" Then
		Response.Write "<tr><td valign='top'></td><td colspan = '2'>Paid By Wire Transfer - Confirmation #" & ChequeNumber & "<br><br></td></tr>"
	    Else
		Response.Write "<tr><td valign='top'></td><td colspan = '2'>Paid By Cheque #" & ChequeNumber & "<br><br></td></tr>"
	    End If
	End If 
    Else %>
  <tr>
    <td width="620" colspan="3" valign="top">
    <b>Terms of Payment</b>
    </td>
  </tr>
  <tr>
    <td valign="top">
    </td>
    <td colspan="2" valign="top">
    <% If CardType = "Wire transfer" Then %>
    No shipments are made until wire transfer is received. Please wire funds in <b>US dollars</b> to: <br><br>
			    <b>Correspondent bank:</b> <br>
			    Bank of America, NY: ABA 026009593 <br><br>
			    <b>Beneficiary bank:</b> <br>
			    TD Bank Swift # TDOMCATTTOR <br>
			    Branch transit #: 2504 <br>
			    Account #: 0460-7300076 <br>
			    Account name: Computational Hydraulics Inc. <br><br>
			    <b>Proof of payment:</b><br>Fax proof of payment to +15194890695 or email
			    <a href="mailto:info@chiwater.com?subject=Wire transfer">info@chiwater.com</a></b>
    <% Else %>
    	Payment due immediately on receipt.<br>
	To pay by credit card, please go to:<br><a href=http://www.chiwater.com/pay.asp<% Response.write "?Item=" & CurrentTransaction & "&Name=" & ClientLastName %>>http://www.chiwater.com/pay.asp<% Response.write "?Item=" & CurrentTransaction & "&Name=" & ClientLastName %></a>
    <% End If %>
    <br><br></td>
  </tr>
  <% End If   
    
    
  %>
</table>
<table width="96%">
  <tr>
    <td width="8%" valign="top"><b>Item</b></td>
    <td width="57%" valign="top" ><b>Description</b></td>
    <td width="15%" align="right" valign="top">
      <b>Unit Price</b>
    </td>
    <td width="10%" align="right" valign="top"><b>Quantity</b></td>
    <td width="10%" align="right" valign="top"><b>Amount</b></td>
  </tr>
  <%
  	For t = 1 to NumCartItems
  		TotalItemCost = CartItemCost(t) * CartItemQuantity(t)
  %>
  <tr>
    <td width="60" valign="top">
      
      <% response.write CartItem(t) %></td>
    <td width="365" valign="top">
     
      <% response.write CartItemDescription(t) %></td>
    <td width="70" align="right" valign="top">
     
      <% response.write FormatCurrency(CartItemCost(t),2) %></td>
    <td width="70" align="right" valign="top">
     
      <% response.write CartItemQuantity(t) %></td>
    <td width="70" align="right" valign="top">
      
      	<% response.write FormatCurrency(TotalItemCost ,2) %>
	</td>
  </tr>
<%
  	Next
%>
  <tr>

    <td width="70" align="right" valign="top" colspan="4">
    <b>SubTotal:</b></td>
    <td width="70" align="right" valign="top">
      
      <% response.write FormatCurrency(SubTotalCost,2) %></td>
  </tr>
  <% If ShippingCost > 0 Then %>
  <tr>

    <td width="140" align="right" valign="top"  colspan="4">
    <b>
      Shipping & handling:</b>
    </td>
    <td width="70" align="right" valign="top">
      
      <% response.write FormatCurrency(ShippingCost,2) %></td>
  </tr>
  <%
	End If 
  		If GSTCost > 0 Then
  %>
  <tr>

    <td width="140" align="right" valign="top" colspan="4">
    <b>GST (#119880979):</b></td>
    <td width="70" align="right" valign="top">
      
      <% response.write FormatCurrency(GSTCost,2) %></td>
  </tr>
  <%
  		End If
		If PSTCost > 0 Then
  %>
  <tr>

    <td width="70" align="right" valign="top" colspan="4">
    <b>
      PST:</b></td>
    <td width="70" align="right" valign="top">
      
      <% response.write FormatCurrency(PSTCost,2) %></td>
  </tr>
  <%
  		End If
		If HSTCost > 0 Then
  %>
  <tr>
    <td width="70" align="right" valign="top" colspan="4">
    <b>
      HST:</b></td>
    <td width="70" align="right" valign="top">
     
      <% response.write FormatCurrency(HSTCost,2) %></td>
  </tr>
  <%
  		End If
		
If IsDate(PaidDate) Then
%>

  <tr>

    <td valign="top" align="right" colspan="4">
    <b>
      Total paid in <% Response.Write CurrencyAbrev %> dollars:</b>
      </td>
    <td width="70" align="right" valign="top">
      <b>
      <% response.write FormatCurrency(TotalCost,2) %></b></td>
  </tr>


<tr><td colspan = '5' align="center"><br><br>
<b>Paid in full with thanks.</b><br>
If you have any questions concerning this statement, call: (519) 767-0197<br>
<br>
<b>THANK YOU FOR YOUR BUSINESS!</b>
</td></tr>

<% Else %>
  <tr>

    <td width="505" align="right" valign="top" colspan="4">
    <b>
      Total due in <% Response.Write CurrencyAbrev %> dollars:</b>
      </td>
    <td width="70" align="right" valign="top">
      <b>
      <% response.write FormatCurrency(TotalCost,2) %></b></td>
  </tr>


<tr><td colspan = '5' align="center"><br><br>
Make all checks payable in <% Response.Write CurrencyAbrev %> dollars to: Computational Hydraulics Int.<br>
If you have any questions concerning this statement, call: (519) 767-0197<br><br>
<b>THANK YOU FOR YOUR BUSINESS!</b>
</td></tr>
<% End If %>
</table>
<br>
</div>
<div class="hidethis">
<script language="javascript">
    document.write("<table width = '97%'><tr>");
    document.write("<td colspan='5' align='right'><span id='fooBar'>&nbsp;</span><br><br></td>");
    document.write("</tr></table>");
</script>
</div>
</form>
					
                      <li class="maincontentbottom_large"></li>
</ul>	                   
<div class="hidethis">
   <p class="p_no_margin">&nbsp;</p>
        <ul>
	<li class="maincontenttop_large"></li>
 <li class="maincontentarea_large">
    <p>
        <b>Note</b>: You may need to remove the headers and footers that your browser includes
        on the printed page. This can typically be done through the print setup dialog for
        your browser. For best print quality, please use Google Chrome, IE 8+ or Firefox. 
        <asp:Label ID="lblEmail" runat="server" Visible="False"></asp:Label>
    </p>
                      <li class="maincontentbottom_large"></li>
</ul>	
</div>				                 
										
				<!-- #EndEditable -->
</li>
				<li class="between">&nbsp;</li>
				<li class="right_content">
				
					<!-- #BeginEditable "RightBar" -->
<!-- #include file = "styles/leftmenu_cart.asp" -->
					<!-- #EndEditable -->	
			  </li>
				<li class="start_end">&nbsp;</li>
		  </ul>
</div>
		
		<div class="footer">
<div style="float:left">
	Copyright © <%= Year(Date) %> CHI. All rights reserved.
</div>

<div style="float:right;text-align:right">	
	<a href="http://www.chiwater.com/Disclaimer.asp">Disclaimer</a>
	 | 
	<a href="http://www.chiwater.com/PrivacyPolicy.asp">Privacy Policy</a>
</div>
		</div>
		
	</div>
</body>
<!-- #EndTemplate --></html>