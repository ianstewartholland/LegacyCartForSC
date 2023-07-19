
<%
	'check to ensure CartItems exist
	If Session("NumCartItems") > 0 Then
		'everything ok
	Else
		Response.Redirect("cart.asp")
	End If
        
        If Request.QueryString("Action") = "Do" Then
            ' kickout spammers
            If Len(Request.Form("ID")) > 0 Then Response.Redirect("cart.asp")
            
            ' email us and the client this quote
            If Session("TotalCost") > 0 And Len(Session("ClientEmail")) > 5 Then
             
                ClientEmailStr = "<font size='2' face='Arial'>Dear " & Session("ClientName") & ",<br><br>" _
                    & "Thank you for your interest in our products and services.<br><br>" & VBCrLf _
                    & "Below is a copy of the quote you generated today. Please note that since the quote was user-generated, there may be a chance of errors in the items or pricing. CHI reserves the right to correct any mistakes before honouring this quote.<br><br>" & VBCrLf _
                    & "If you have any questions or concerns, please contact us at 519-767-0197 or info@chiwater.com.<br><br></font>" & VBCrLf & Session("ClientEmailStr")
                
                    On Error Resume Next
                    Set objNewMail = CreateObject("CDO.Message")
                    objNewMail.Subject = "CHI quote"
                    objNewMail.From = "store@chiwater.com"
                    objNewMail.To = Session("ClientEmail")
                    objNewMail.Bcc = "info@chiwater.com"
                    objNewMail.HTMLBody = ClientEmailStr
            
                    objNewMail.Configuration.Fields.Item _
                            ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                    'Name or IP of remote SMTP server
                    objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="localhost"
                    objNewMail.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
                    objNewMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                    
                    objNewMail.Configuration.Fields.Update
            
                    objNewMail.Send
                    Set objNewMail = Nothing
                    
            End If
            
            Response.Redirect("quote.asp")
        End If
'	If Len(Request.QueryString("Transaction")) > 1 Then
'            CurrentTransaction = Request.QueryString("Transaction")
'	Else
'            Response.Redirect "currentorders.asp"
'	End If
'
'    	' set connection object here
'	set conn=Server.CreateObject("ADODB.Connection")
'	conn.Provider="Microsoft.Jet.OLEDB.4.0"
'	conn.Open transcationPath
'
'	'connect to database and load the information for this page
'	Set objRec = Server.CreateObject ("ADODB.Recordset")	
'	'objRec.Open "Transaction", transcationPath, 3, 3, 2
'    objRec.Open "[Transaction]", conn
'	objRec.Filter = "TransactionNumber = '" & CurrentTransaction & "'"
'
'	objRec.MoveFirst

	'load client info ===========================

	'Session("ClientAddress") = "Yes"
	ClientPrefix = Session("ClientPrefix")
	ClientFirstName = Session("ClientFirstName")
	ClientLastName = Session("ClientLastName")
	ClientName =  Trim(ClientPrefix & " " & Trim(ClientFirstName & " " & ClientLastName))
	ClientCompany = Session("ClientCompany")
	ClientStreetAddress = Session("ClientStreetAddress")
	ClientStreetAddress2 = Session("ClientStreetAddress2")
	ClientCity = Session("ClientCity")
	ClientProvince = Session("ClientProvince")
	ClientCountry = Session("ClientCountry")
	ClientPostalCode = Session("ClientPostalCode")
	ClientEmail = Session("ClientEmail")
	ClientTelephone = Session("ClientTelephone")
	ClientFax = Session("ClientFax")
	ClientWebSite = Session("ClientWebSite")
	ClientType = Session("ClientType")
	'PreviousLicensee = ObjRec("PreviousLicensee")

	'load shipping info ===========================

	'Session("ShipAddress") = "Yes"
	ShipPrefix = Session("ShipPrefix")
	ShipFirstName = Session("ShipFirstName")
	ShipLastName = Session("ShipLastName")
	ShipName =  Trim(ShipPrefix & " " & Trim(ShipFirstName & " " & ShipLastName))
	ShipCompany = Session("ShipCompany")
	ShipStreetAddress = Session("ShipStreetAddress")
	ShipStreetAddress2 = Session("ShipStreetAddress2")
	ShipCity = Session("ShipCity")
	ShipProvince = Session("ShipProvince")
	ShipCountry = Session("ShipCountry")
	ShipPostalCode = Session("ShipPostalCode")
	ShipEmail = Session("ShipEmail")
	ShipTelephone = Session("ShipTelephone")
	ShipFax = Session("ShipFax")
	
	'load billing info ===========================

	'Session("BillAddress") = "Yes"
	BillPrefix = Session("BillPrefix")
	BillFirstName = Session("BillFirstName")
	BillLastName = Session("BillLastName")
	BillName =  Trim(BillPrefix & " " & Trim(BillFirstName & " " & BillLastName))
	BillCompany = Session("BillCompany")
	BillStreetAddress = Session("BillStreetAddress")
	BillStreetAddress2 = Session("BillStreetAddress2")
	BillCity = Session("BillCity")
	BillProvince = Session("BillProvince")
	BillCountry = Session("BillCountry")
	BillPostalCode = Session("BillPostalCode")
	BillEmail = Session("BillEmail")
	BillTelephone = Session("BillTelephone")
	BillFax = Session("BillFax")

	'load currencies ======================================
	
	CurrencyStr = Session("Currency")
	CurrencyAbrev = Session("CurrencyAbrev")
	'If IsNull(Session("CurrencyExchange")) Then
	'    CurrencyExchange = -1 ' put -1 in case the currency has not been set, people would be able to tell
	'Else
	'    CurrencyExchange = ObjRec("CurrencyExchange")
	'End If
	
'	CurrencyExchange = 0 ' assign default value
'	CurrencyExchange = ObjRec("CurrencyExchange")
'        If IsNull(CurrencyExchange) Then CurrencyExchange = 0
	
	'ShippingCalculated = "Yes"
	ShippingCost = Session("ShippingCost")
        If IsNull(ShippingCost) Then ShippingCost = 0

	'TaxCalculated = "Yes"
        HSTCost = Session("HSTCost")
        If IsNull(HSTCost) Then HSTCost = 0
	GSTCost = Session("GSTCost")
        If IsNull(GSTCost) Then GSTCost = 0
	PSTCost = Session("PSTCost")
        If IsNull(PSTCost) Then PSTCost = 0

	'Session("SubTotalCalculated") = "Yes"
	SubTotalCost = Session("SubTotalCost")
        If IsNull(SubTotalCost) Then SubTotalCost = 0
	TotalCost = Session("TotalCost")
        If IsNull(TotalCost) Then TotalCost = 0
'	OriginalTotalCost = Session("OriginalTotalCost")
'        If IsNull(OriginalTotalCost) Then OriginalTotalCost = 0

	'load dates ===========================================

	'OrderDate = ObjRec("OrderDate")
        OrderDate = Date
	PurchaseOrder = Session("PurchaseOrder")

	
	
	'objRec.Close
	'Set objRec = Nothing	

%>
<!DOCTYPE html>
<html>

<head>
<title>CHI Quote #<%= CurrentTransaction %></title>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
	<link href="styles/style2.css" rel="STYLESHEET" type="text/css" media="screen" />
	<link href="styles/print.css" rel="STYLESHEET" type="text/css" media="print" />
	<link rel="shortcut icon" href="images/chilogo.ico" />
   <style type="text/css">
        .tablecells
        {
           display:none;
           visibility:hidden;
        }
    </style>
    <script type="text/javascript">
    function redirect()
    {
        var txt = document.getElementById('ID');

        if (txt != null)
        {
            if (txt.value.length > 0)
                window.location.href = "cart.asp";
            else
                window.location.href = "quote.asp?Action=Do";
        }
    }
    </script>
        <script async src="https://www.googletagmanager.com/gtag/js?id=UA-8955620-8"></script>
    <script>
        window.dataLayer = window.dataLayer || [];
        function gtag(){dataLayer.push(arguments);}
        gtag('js', new Date());

        gtag('config', 'UA-8955620-8');
    </script>
</head>

<body>
<form>
    <table border='0' width='620' class='hiddenheader'>    
        <tr>
            <td align='right'>
                <input type="button" value="Email me the quote" onclick="redirect();" />&nbsp;
                <input type="button" value="Print this page" onclick="window.print();return false;" />
            </td>
        </tr>
    </table>
</form>
<%
Session("ClientEmailStr") = ""
ClientEmailStr = "<table border='0' width='620' cellpadding='0' cellspacing='0'>" & VBCrLf
ClientEmailStr = ClientEmailStr & "<tr>"
ClientEmailStr = ClientEmailStr & "    <td width='370' valign='top' colspan='2'>"
ClientEmailStr = ClientEmailStr & "        <font face='Arial' size='5'><b>Computational Hydraulics Int.</b></font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "    <td width='250' align='right' valign='top'>"
ClientEmailStr = ClientEmailStr & "    <font face='Arial' size='5'><b>Unofficial Quote</b></font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "  </tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & "  <tr>"
ClientEmailStr = ClientEmailStr & "    <td width='50' valign='top'>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "    <td width='270' valign='top'>"
ClientEmailStr = ClientEmailStr & "        <font face='Arial' size='2'>"
ClientEmailStr = ClientEmailStr & "        147 Wyndham Street, Suite 202<br>"
ClientEmailStr = ClientEmailStr & "	Guelph, Ontario, Canada, N1H 4E9<br>"
ClientEmailStr = ClientEmailStr & "        Tel: (519) 767-0197 Fax: (519) 489-0695<br>"
ClientEmailStr = ClientEmailStr & "        Email: <a href='mailto:info@chiwater.com'>info@chiwater.com</a><br>"
ClientEmailStr = ClientEmailStr & "        Web: www.chiwater.com</font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "    <td width='250' align='right' valign='top'>"
ClientEmailStr = ClientEmailStr & "    <font face='Arial' size='2'>" & FormatDateTime(OrderDate,1) & "</font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "  </tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & "</table>" & VBCrLf

ClientEmailStr = ClientEmailStr & "<br>"

ClientEmailStr = ClientEmailStr & "<table border='0' width='620'>" & VBCrLf
ClientEmailStr = ClientEmailStr & "  <tr>"
ClientEmailStr = ClientEmailStr & "    <td width='620' colspan='3' valign='top'>"
ClientEmailStr = ClientEmailStr & "    <font face='Arial' size='2'><b>"
ClientEmailStr = ClientEmailStr & "    Client Information</b></font>"
ClientEmailStr = ClientEmailStr & "    </td>"
ClientEmailStr = ClientEmailStr & "  </tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & "  <tr>"
ClientEmailStr = ClientEmailStr & "    <td width='50' valign='top'>"
ClientEmailStr = ClientEmailStr & "    </td>"
ClientEmailStr = ClientEmailStr & "    <td width='270' valign='top'><font size='2' face='Arial'>"

    	ClientEmailStr = ClientEmailStr &  ClientName & "<br>"
    	ClientEmailStr = ClientEmailStr &  ClientCompany & "<br>"
    	ClientEmailStr = ClientEmailStr &  ClientStreetAddress & "<br>"
     	If ClientStreetAddress2 <> "" Then ClientEmailStr = ClientEmailStr &  ClientStreetAddress2 & "<br>"
    	ClientEmailStr = ClientEmailStr &  ClientCity & " " & ClientProvince & " " & ClientPostalCode & "<br>"
    	ClientEmailStr = ClientEmailStr &  ClientCountry & "<br><br>"

ClientEmailStr = ClientEmailStr & "        </font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "    <td width='300' valign='top'><font size='2' face='Arial'>"

    	ClientEmailStr = ClientEmailStr & "Email: " & ClientEmail & "</a><br>"
    	ClientEmailStr = ClientEmailStr & "Tel: " & ClientTelephone & "<br>"
    	ClientEmailStr = ClientEmailStr & "Fax: " & ClientFax & "<br><br>"

ClientEmailStr = ClientEmailStr & "</font>"
ClientEmailStr = ClientEmailStr & "    </td>"
ClientEmailStr = ClientEmailStr & "  </tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & "  <tr>"
ClientEmailStr = ClientEmailStr & "    <td width='620' colspan='3' valign='top'>"
ClientEmailStr = ClientEmailStr & "    <font face='Arial' size='2'><b>"
ClientEmailStr = ClientEmailStr & "    Shipping Address</b></font>"
ClientEmailStr = ClientEmailStr & "    </td>"
ClientEmailStr = ClientEmailStr & "  </tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & " <tr>"
ClientEmailStr = ClientEmailStr & "    <td width='50' valign='top'>"
ClientEmailStr = ClientEmailStr & "    </td>"

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

ClientEmailStr = ClientEmailStr & "    <td width='570' colspan='2' valign='top'><font size='2' face='Arial'>"
ClientEmailStr = ClientEmailStr & "    Same as client address above...<br><br>"
ClientEmailStr = ClientEmailStr & "    </font>"
ClientEmailStr = ClientEmailStr & "    </td>"

    Else

ClientEmailStr = ClientEmailStr & "    <td width='270' valign='top'><font size='2' face='Arial'>"

    	ClientEmailStr = ClientEmailStr & ShipName & "<br>"
    	ClientEmailStr = ClientEmailStr & ShipCompany & "<br>"
    	ClientEmailStr = ClientEmailStr & ShipStreetAddress & "<br>"
    	If ShipStreetAddress2 <> "" Then ClientEmailStr = ClientEmailStr & ShipStreetAddress2 & "<br>"
    	ClientEmailStr = ClientEmailStr & ShipCity & " " & ShipProvince & " " & ShipPostalCode & "<br>"
    	ClientEmailStr = ClientEmailStr & ShipCountry & "<br><br>"

ClientEmailStr = ClientEmailStr & "        </font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "    <td width='300' valign='top'><font size='2' face='Arial'>"

    	ClientEmailStr = ClientEmailStr & "Email: " & ShipEmail & "<br>"
    	ClientEmailStr = ClientEmailStr & "Tel: " & ShipTelephone & "<br>"
    	ClientEmailStr = ClientEmailStr & "Fax: " & ShipFax & "<br><br>"

ClientEmailStr = ClientEmailStr & "</font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf

    End If

ClientEmailStr = ClientEmailStr & "  </tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & "  <tr>"
ClientEmailStr = ClientEmailStr & "    <td width='620' colspan='3' valign='top'>"
ClientEmailStr = ClientEmailStr & "    <font face='Arial' size='2'><b>"
ClientEmailStr = ClientEmailStr & "    Billing Address</b></font>"
ClientEmailStr = ClientEmailStr & "    </td>"
ClientEmailStr = ClientEmailStr & "  </tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & "  <tr>"
ClientEmailStr = ClientEmailStr & "    <td width='50' valign='top'>"
ClientEmailStr = ClientEmailStr & "    </td>"

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

ClientEmailStr = ClientEmailStr & "    <td width='570' colspan='2' valign='top'><font size='2' face='Arial'>"
ClientEmailStr = ClientEmailStr & "    Same as client address above...<br><br>"
ClientEmailStr = ClientEmailStr & "    </font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf

    Else

ClientEmailStr = ClientEmailStr & "    <td width='270' valign='top'><font size='2' face='Arial'>"

    	ClientEmailStr = ClientEmailStr & BillName & "<br>"
     	If BillCompany <> "" Then ClientEmailStr = ClientEmailStr & BillCompany & "<br>"
    	ClientEmailStr = ClientEmailStr & BillStreetAddress & "<br>"
    	If BillStreetAddress2 <> "" Then ClientEmailStr = ClientEmailStr & BillStreetAddress2 & "<br>"
    	ClientEmailStr = ClientEmailStr & BillCity & " " & BillProvince & " " & BillPostalCode & "<br>"
    	ClientEmailStr = ClientEmailStr & BillCountry & "<br><br>"

ClientEmailStr = ClientEmailStr & "         </font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "    <td width='300' valign='top'><font size='2' face='Arial'>"

    	ClientEmailStr = ClientEmailStr & "Email: " & BillEmail & "<br>"
    	ClientEmailStr = ClientEmailStr & "Tel: " & BillTelephone & "<br>"
    	ClientEmailStr = ClientEmailStr & "Fax: " & BillFax & "<br><br>"

ClientEmailStr = ClientEmailStr & "        </font>"
ClientEmailStr = ClientEmailStr & "    </td>" & VBCrLf

    End If

ClientEmailStr = ClientEmailStr & "  </tr>"

    If PurchaseOrder <> "" Then

ClientEmailStr = ClientEmailStr & "		<tr>"
ClientEmailStr = ClientEmailStr & "			<td width='620' colspan='3' valign='top'><font face='Arial' size='2'><b>Purchase Order #</b></font></td>"
ClientEmailStr = ClientEmailStr & "		</tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & "		<tr>"
ClientEmailStr = ClientEmailStr & "			<td width='50' valign='top'></td>"
ClientEmailStr = ClientEmailStr & "			<td width='270' valign='top'><font size='2' face='Arial'>" & "<br><br>" &"</font></td>"
ClientEmailStr = ClientEmailStr & "			<td width='300' valign='top'></td>"
ClientEmailStr = ClientEmailStr & "		</tr>" & VBCrLf

    End If

ClientEmailStr = ClientEmailStr & "</table>" & VBCrLf


'Response.Write ClientEmailStr & "</body></html>"

ClientEmailStr = ClientEmailStr & "<table border='0' width='620'>"
ClientEmailStr = ClientEmailStr & "    <tr>"
ClientEmailStr = ClientEmailStr & "        <td width='8%'><font size='2' face='Arial'><b>Item</b></font></td>"
ClientEmailStr = ClientEmailStr & "        <td width='360'><font size='2' face='Arial'><b>Description</b></font></td>"
ClientEmailStr = ClientEmailStr & "        <td width='90' align='right'><font size='2' face='Arial'><b>Unit Price</b></font></td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "        <td width='70' align='right'><font size='2' face='Arial'><b>Quantity</b></font></td>"
ClientEmailStr = ClientEmailStr & "        <td width='70' align='right'><font size='2' face='Arial'><b>Amount</b></font></td>"
ClientEmailStr = ClientEmailStr & "    </tr>" & VBCrLf

        ' & vbTab
        'EmailStr = ""
        
          For t = 1 to Session("NumCartItems")
            'If Session("TotalCost") > 0 Then
            '    EmailStr = Session("CartItemDescription(" & t & ")") & " (" & Session("CartItem(" & t & ")") & ")" & vbCrLf & vbcrlf & EmailStr
            'End If
                
            TotalItemCost = (Session("CartItemCost(" & t & ")") * Session("CartItemQuantity(" & t & ")"))

ClientEmailStr = ClientEmailStr & "    <tr>"
ClientEmailStr = ClientEmailStr & "        <td valign='top'><font size='2' face='Arial'>" & Session("CartItem(" & t & ")") & "</font></td>"
ClientEmailStr = ClientEmailStr & "        <td><font size='2' face='Arial'>" & Session("CartItemDescription(" & t & ")") & "</font></td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "        <td align='right'><font size='2' face='Arial'>" & FormatCurrency(Session("CartItemCost(" & t & ")"),2) & "</font></td>"
ClientEmailStr = ClientEmailStr & "        <td align='right'><font size='2' face='Arial'>" & Session("CartItemQuantity(" & t & ")") & "</font></td>"
 ClientEmailStr = ClientEmailStr & "       <td align='right'><font size='2' face='Arial'>" & FormatCurrency(TotalItemCost,2) & "</font></td>"
ClientEmailStr = ClientEmailStr & "    </tr>" & VBCrLf

  		Next
                

ClientEmailStr = ClientEmailStr & "    <tr>"
ClientEmailStr = ClientEmailStr & "        <td align='right' colspan='4'><font size='2' face='Arial'><b>SubTotal:</b></font></td>"
ClientEmailStr = ClientEmailStr & "        <td align='right'><font size='2' face='Arial'><b>" & FormatCurrency(Session("SubTotalCost"),2) & "</b></font></td>"
ClientEmailStr = ClientEmailStr & "    </tr>" & VBCrLf

	If Session("ShipCountry") <> "" And Session("ShipProvince") <> "" Or Session("PayQuoteMode") = "Yes" Then
  		If Session("ShippingCost") > 0 Then
ClientEmailStr = ClientEmailStr & "    <tr>"
ClientEmailStr = ClientEmailStr & "        <td align='right' colspan='4'><font size='2' face='Arial'><b>Shipping:</b></font></td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "        <td align='right'><font size='2' face='Arial'><b>" & FormatCurrency(Session("ShippingCost"),2) & "</b></font></td>"
ClientEmailStr = ClientEmailStr & "  </tr>" & VBCrLf

		End If
  		If Session("HSTCost") > 0 Then
ClientEmailStr = ClientEmailStr & "    <tr>"
ClientEmailStr = ClientEmailStr & "        <td align='right' colspan='4'><font size='2' face='Arial'><b>HST (#119880979):</b></font></td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "        <td align='right'><font size='2' face='Arial'><b>" & FormatCurrency(Session("HSTCost"),2) & "</b></font></td>"
ClientEmailStr = ClientEmailStr & "    </tr>" & VBCrLf

                End If
  		If Session("GSTCost") > 0 Then

ClientEmailStr = ClientEmailStr & "    <tr>"
ClientEmailStr = ClientEmailStr & "        <td align='right' colspan='4'><font size='2' face='Arial'><b>GST (#119880979):</b></font></td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "        <td align='right'><font size='2' face='Arial'><b>" & FormatCurrency(Session("GSTCost"),2) & "</b></font></td>"
ClientEmailStr = ClientEmailStr & "    </tr>" & VBCrLf

  		End If
 
ClientEmailStr = ClientEmailStr & "    <tr>"
ClientEmailStr = ClientEmailStr & "        <td align='right' colspan='4'><font size='2' face='Arial'><b>Total (" & Session("CurrencyAbrev") & "$):</b></font></td>" & VBCrLf
ClientEmailStr = ClientEmailStr & "        <td align='right'><font size='2' face='Arial'><b>" & FormatCurrency(Session("TotalCost"),2) & "</b></font></td>"
ClientEmailStr = ClientEmailStr & "    </tr>" & VBCrLf
      End If 
ClientEmailStr = ClientEmailStr & "</table>"

ClientEmailStr = ClientEmailStr & "<table border='0' width='620'>"
ClientEmailStr = ClientEmailStr & "    <tr>"
ClientEmailStr = ClientEmailStr & "	<td><center>"
ClientEmailStr = ClientEmailStr & "<font size='2' face='Arial'><br>This quote was user-generated without review by CHI staff. CHI reserves the right to correct any mistakes.<br><br>If you have any questions concerning this quote, call: (519) 767-0197<br><br>"
ClientEmailStr = ClientEmailStr & "<b>THANK YOU FOR YOUR BUSINESS!</b></font>"
ClientEmailStr = ClientEmailStr & "	</center></td>"
ClientEmailStr = ClientEmailStr & "    </tr>" & VBCrLf
ClientEmailStr = ClientEmailStr & "</table>" 

Response.Write ClientEmailStr

Session("ClientEmailStr") = ClientEmailStr
%>
</body>
</html>


 