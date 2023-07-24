<!--#include file="LogFile.asp"-->
<!--#include virtual ="/chi/includes/printxml.asp"-->
<!--#include virtual = "/chi/includes/sendemail.asp"-->
<!--#include virtual ="/chi/includes/getexchangerate.asp"-->
<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<!--#include virtual ="/chi/includes/jsonObject.class.asp"-->

<%
'
' on May 24, 2011 we got three orders with items but empty client information
' this is to add a check to prevent similar case happening again
'
If (Session("PayQuoteMode") <> "Yes" And Session("ShowClient") And Session("CardType") <> "") Or (Session("PayQuoteMode") = "Yes" And Session("CardType") <> "") Or (Session("TotalCost") = 0) Then
    ' good to go
Else
    Session("ErrMsg") = "You should double check the information you entered for your payment and client information."
    Response.Redirect "paymentform.asp"
End if
	'check session is still active
	If Session("NumCartItems") <= 0 Then
		Response.Redirect "cart.asp"
	End If

	dim ItemStr
	dim EmailStr
	dim EmailTitleStr
	dim NewTransactionNumber
	dim t
	dim objConn
	dim objRec
	dim objNewMail
	dim cardAuthNo
	payOrder = False

	On Error Resume Next

	NewTransactionNumber = GetTransactionNumber()


	' Turn pivotal payment processing on/off.
	Pivotal = 1

	' Reset payment.
	Session("TotalPaid") = 0



	'init email string
	EmailStr = Session("ClientFirstName") & " " &  Session("ClientLastName") & vbCrLf
	EmailStr = EmailStr & Session("ClientCompany") & vbCrLf
	EmailStr = EmailStr & Session("ClientStreetAddress") & vbcrlf
	If Session("ClientStreetAddress2") <> "" Then EmailStr = EmailStr & Session("ClientStreetAddress2") & vbcrlf
	EmailStr = EmailStr & Session("ClientCity") & ", " & Session("ClientProvince") & ", " & Session("ClientPostalCode") & vbcrlf
	EmailStr = EmailStr & Session("ClientCountry") & vbcrlf
	EmailStr = EmailStr & "Email: " & Session("ClientEmail") & vbcrlf
	EmailStr = EmailStr & "Tel: " & Session("ClientTelephone") & vbcrlf
	EmailStr = EmailStr & "Fax: " & Session("ClientFax") & vbcrlf & vbcrlf

	'create item string
	ItemStr = ""
	isTrial = False
	isWebinar = False
	workshopItem = ""
    hasMultipleItem = False
	For t = 1 to Session("NumCartItems")
		If Session("CartItem(" & t & ")") = "S212" Then isTrial = True
		If Left(Session("CartItem(" & t & ")"), 1) = "P" Then isWebinar = True

		If Left(Session("CartItem(" & t & ")"), 1) = "W" Then
                    workshopItem = Session("CartItem(" & t & ")")
                    EmailStr = Replace(Replace(Session("CartItemDescription(" & t & ")"), "workshop & time limited software", "time limited software with workshop " & Session("CartItem(" & t & ")")), " & ", "-") & vbCrLf & vbcrlf & EmailStr
                Else
                    EmailStr = Session("CartItemDescription(" & t & ")") & " (" & Session("CartItem(" & t & ")") & ")" & vbCrLf & vbcrlf & EmailStr
                End If

		    ItemStr = ItemStr & Session("CartItem(" & t & ")") & "/" & Session("CartItemDescription(" & t & ")") & "/" & Session("CartItemCost(" & t & ")") & "/" & Session("CartItemQuantity(" & t & ")") & "/"

            If Session("CartItemQuantity(" & t & ")")  > 1 Then hasMultipleItem = True
	Next

If LCASE(Session("ClientEmail")) = "linbjorem@gmail.com" Then
    Session("ErrMsg") = "Please double-check your credit card information and try again. Most often, it's something as simple as a typo or an old expiration date that's causing the payment to fail. Don't use spaces or dashes in your card number. If everything seems correct and the charge still won't go through, your card issuer may be blocking the payment. Card-issuing banks have many security measures in place for online purchases and subscriptions, especially when a card is being used for the first time." & vbCrLf

	'send them back to CC info
	Response.Redirect "paymentform.asp"
End If


'check for credit card payments
If (UCase(Session("CardType")) = "VISA" Or _
				UCase(Session("CardType")) = "MASTERCARD" Or _
				UCase(Session("CardType")) = "AMERICAN EXPRESS") And _
				Pivotal = 1 Then


				'begin pivotal implementation -------------------------------

				GatewayID = "5447"
				UserName = "chiw4317"
				Password = "Rocky@34"

				' Remove any non-numeric number from the credit card number and the expiry date.
				strCardNum = MakeNumeric(Session("CardNumber"))
				strExpDate = MakeNumeric(Session("CardExpiry"))
				strCCName = Session("CardName")
				strAmount = Session("TotalCost")

				Set XmlHTTP = CreateObject("MSXML2.ServerXMLHTTP")
				Set myXML =Server.CreateObject("MSXML.DOMDocument")
				myXML.Async=False
				serviceURL = "https://secure1.pivotalpayments.com/smartpayments/transact.asmx?op=ProcessCreditCard"

				' branch for CAD or USD account based on the currency
				If Session("CurrencyAbrev") = "CAD" Then
				    SOAPParameters = "<UserName>" & UserName & "</UserName>"
				ElseIf Session("CurrencyAbrev") = "US" Then
				    SOAPParameters = "<UserName>comp9718</UserName>"
				End If

				SOAPParameters = SOAPParameters & " <Password>" & Password & "</Password>"
				SOAPParameters = SOAPParameters & " <TransType>Sale</TransType>"
				SOAPParameters = SOAPParameters & " <CardNum>" & strCardNum & "</CardNum>"
				SOAPParameters = SOAPParameters & " <ExpDate>" & strExpDate & "</ExpDate>"
				SOAPParameters = SOAPParameters & " <MagData>" & "</MagData>"
				SOAPParameters = SOAPParameters & " <NameOnCard>" & strCCName & "</NameOnCard>"
				SOAPParameters = SOAPParameters & " <Amount>" & strAmount & "</Amount>"
				SOAPParameters = SOAPParameters & " <InvNum>" & NewTransactionNumber & "</InvNum>"
				SOAPParameters = SOAPParameters & " <PNRef>" & "</PNRef>"
				SOAPParameters = SOAPParameters & " <Zip>" & "</Zip>"
				SOAPParameters = SOAPParameters & " <Street>" & "</Street>"
				SOAPParameters = SOAPParameters & " <CVNum>" & "</CVNum>"
				SOAPParameters = SOAPParameters & " <ExtData>" & "</ExtData>"

				XmlHTTP.Open "POST", serviceUrl, False

				XmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
				XmlHTTP.setRequestHeader "SOAPAction", "http://TPISoft.com/SmartPayments/ProcessCreditCard"
				SOAPRequest = "<?xml version='1.0' encoding='utf-8'?> <soap:Envelope"
				SOAPRequest = SOAPRequest & " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"""
				SOAPRequest = SOAPRequest & " xmlns:xsd=""http://www.w3.org/2001/XMLSchema"""
				SOAPRequest = SOAPRequest & " xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
				SOAPRequest = SOAPRequest & " <soap:Body>"
				SOAPRequest = SOAPRequest & " <ProcessCreditCard xmlns=""http://TPISoft.com/SmartPayments/"">"
				SOAPRequest = SOAPRequest & SOAPParameters
				SOAPRequest = SOAPRequest & " </ProcessCreditCard>"
				SOAPRequest = SOAPRequest & " </soap:Body>"
				SOAPRequest = SOAPRequest & " </soap:Envelope>"

				XmlHTTP.send SOAPRequest
				myXML.load(XmlHTTP.responseXML)

				Set nodesResult=myXML.documentElement.selectNodes("//Result")
				Set nodesRespMSG=myXML.documentElement.selectNodes("//RespMSG")
				Set nodesMessage=myXML.documentElement.selectNodes("//Message")
				Set nodesMessage1=myXML.documentElement.selectNodes("//Message1")
				Set nodesAuthCode=myXML.documentElement.selectNodes("//AuthCode")
				Set nodesPNRef=myXML.documentElement.selectNodes("//PNRef")
				Set nodesHostCode=myXML.documentElement.selectNodes("//HostCode")
				Set nodesHostURL=myXML.documentElement.selectNodes("//HostURL")
				Set nodesReceiptURL=myXML.documentElement.selectNodes("//ReceiptURL")
				Set nodesGetAVSResult=myXML.documentElement.selectNodes("//GetAVSResult")
				Set nodesGetAVSResultTXT=myXML.documentElement.selectNodes("//GetAVSResultTXT")
				Set nodesGetStreetMatchTXT=myXML.documentElement.selectNodes("//GetStreetMatchTXT")
				Set nodesGetZipMatchTXT=myXML.documentElement.selectNodes("//GetZipMatchTXT")
				Set nodesGetCVResult=myXML.documentElement.selectNodes("//GetCVResult")
				Set nodesGetCVResultTXT=myXML.documentElement.selectNodes("//GetCVResultTXT")
				Set nodesGetGetOrigResult=myXML.documentElement.selectNodes("//GetGetOrigResult")
				Set nodesGetCommercialCard=myXML.documentElement.selectNodes("//GetCommercialCard")
				Set nodesWorkingKey=myXML.documentElement.selectNodes("//WorkingKey")
				Set nodesKeyPointer=myXML.documentElement.selectNodes("//KeyPointer")
				Set nodesExtData=myXML.documentElement.selectNodes("//ExtData")

				If nodesResult(0).text <> "0" Then
								' Send email with error information.
								EmailErr = "Failed credit card submission for order from client: " & vbCrLf & vbCrLf
								EmailErr = EmailErr & Session("ClientFirstName") & " " &  Session("ClientLastName") & vbCrLf
								EmailErr = EmailErr & Session("ClientCompany") & vbCrLf
								EmailErr = EmailErr & Session("ClientCity") & vbCrLf
								EmailErr = EmailErr & Session("ClientCountry") & vbCrLf
								EmailErr = EmailErr & "Email: " & Session("ClientEmail") & vbCrLf
								EmailErr = EmailErr & "Tel: " & Session("ClientTelephone") & vbCrLf
								EmailErr = EmailErr & "Fax: " & Session("ClientFax") & vbCrLf & vbCrLf
								If Session("ClientNotes") <> "" Then
												EmailErr = EmailErr & vbCrLf & vbCrLf & "Client notes:" & vbCrLf & Session("ClientNotes") & vbCrLf & vbCrLf
								End If

								EmailErr = EmailErr & "Client tried to order: " & ItemStr
								EmailErr = EmailErr & vbCrLf & vbCrLf & "Total: " & FormatCurrency(Session("TotalCost"),2) & vbCrLf & vbCrLf

								If True Then
												' Send the xml response back as well
												EmailErr = EmailErr & "-------------------------------------" & vbCrLf
												EmailErr = EmailErr & XmlHTTP.responseXML.xml & vbCrLf
												EmailErr = EmailErr & "-------------------------------------" & vbCrLf
								End If

								WritePivotalLog("Error Processing Credit Card")
								WritePivotalLog(DisplayNode(XmlHTTP.responseXML.childNodes))

		SendMail "meghan@chiwater.com; sharon@chiwater.com", _
				"store@chiwater.com", _
				"Failed transaction", _
				EmailErr


								'set the error msg
								'Session("ErrMsg") = "There was an error in processing your credit card. Please try again. Please check the information provided and ensure it is correct. Email us if problems persist." & vbCrLf
                                Session("ErrMsg") = "Please double-check your credit card information and try again. Most often, it's something as simple as a typo or an old expiration date that's causing the payment to fail. Don't use spaces or dashes in your card number. If everything seems correct and the charge still won't go through, your card issuer may be blocking the payment. Card-issuing banks have many security measures in place for online purchases and subscriptions, especially when a card is being used for the first time." & vbCrLf

								'send them back to CC info
								Response.Redirect "paymentform.asp"
				End If

				WritePivotalLog("Successfully processed credit card.")
				WritePivotalLog(DisplayNode(XmlHTTP.responseXML.childNodes))
				Session("TotalPaid") = Session("TotalCost")
				Session("PaidDate") = Date

				' Save the returned reference number
				Session("CardAuthorizationNo") = nodesPNRef(0).Text

				'end pivotal implementation ---------------------------------

End If 'end of pivotal if statement

'below executes as long as:
' - not a free trial -edit: now free trials do go through this
' - if pivotal was turned on, and everything was fine (since they wouldn't have been redirected)
' - if pivotal turned off, then it will always execute for credit cards
Set objRec = Server.CreateObject ("ADODB.Recordset")

	' set connection object here
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider="Microsoft.Jet.OLEDB.4.0"
    ' to test
    'transcationPath = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WWWUsers\LocalUser\chi\computationalhydraulics_com\databases\transactiondb2.mdb"
	conn.Open transcationPath

'strConnect = transcationPath

'objRec.Open "Transaction", strConnect, 3, 3, 2

If Session("PayQuoteMode") = "Yes" Then

	If Session("CardType") <> "Invoice" And Session("CardType") <> "Wire transfer" Then
		objRec.Open "Select * From [Transaction] Where [TransactionNumber] = '" & Session("CurrentTransaction") & "'", conn, 0, 3
		' also mark the quote so the client cannot pay twice
		'objRec.Filter = "TransactionNumber = '" & Session("CurrentTransaction") & "'"
		objRec.MoveFirst

		objRec("TransactionCreated") = Date
		objRec.Update
		objRec.Close
    End If

    ' update notes for the record
    Set objRecNewNotes = Server.CreateObject ("ADODB.Recordset")
    objRecNewNotes.Open "Notes", transcationPath, 3, 3, 2

    If UCase(Left(Session("CurrentTransaction"), 1)) = "Q" Then
	objRecNewNotes.AddNew
	objRecNewNotes("Date") = Date
	objRecNewNotes("RecordID") = Session("CurrentTransaction")
	objRecNewNotes("Content") = "Quote paid by client in transaction #" & NewTransactionNumber
	objRecNewNotes("Author") = "Shopping Cart"
	objRecNewNotes.Update
    End If

    objRecNewNotes.AddNew
    objRecNewNotes("Date") = Date
    objRecNewNotes("RecordID") = NewTransactionNumber
    objRecNewNotes("Content") = "Transaction paid by client for #" & Session("CurrentTransaction")
    objRecNewNotes("Author") = "Shopping Cart"
    objRecNewNotes.Update

    objRecNewNotes.Close
    Set objRecNewNotes = Nothing

    objRec.Open "Transaction", strConnect, 3, 3, 2

End If


If Not payOrder Then
WritePivotalLog("now in pay order mode")
    objRec.Open "[Transaction]", conn, 1, 3
    objRec.AddNew
Else
WritePivotalLog("now in pay quote or normal mode")
    objRec.Open "Select * From [Transaction] Where [TransactionNumber] = '" & NewTransactionNumber & "'", conn, 0, 3
    'objRec.Filter = "TransactionNumber = '" & NewTransactionNumber & "'"
    objRec.MoveFirst
End If

objRec("ClientPrefix") = Session("ClientPrefix")
objRec("ClientFirstName") = Session("ClientFirstName")
objRec("ClientLastName") = Session("ClientLastName")
objRec("ClientCompany") = Session("ClientCompany")
objRec("ClientStreetAddress") = Session("ClientStreetAddress")
objRec("ClientStreetAddress2") = Session("ClientStreetAddress2")
objRec("ClientCity") = Session("ClientCity")
objRec("ClientProvince") = Session("ClientProvince")
objRec("ClientCountry") = Session("ClientCountry")
objRec("ClientPostalCode") = Session("ClientPostalCode")
objRec("ClientEmail") = Session("ClientEmail")
objRec("ClientTelephone") = Session("ClientTelephone")
objRec("ClientFax") = Session("ClientFax")
objRec("ClientWebSite") = Session("ClientWebSite")
objRec("ClientType") = Session("ClientType")
objRec("PreviousLicensee") = Session("PreviousLicensee")

objRec("ShipPrefix") = Session("ShipPrefix")
objRec("ShipFirstName") = Session("ShipFirstName")
objRec("ShipLastName") = Session("ShipLastName")
objRec("ShipCompany") = Session("ShipCompany")
objRec("ShipStreetAddress") = Session("ShipStreetAddress")
objRec("ShipStreetAddress2") = Session("ShipStreetAddress2")
objRec("ShipCity") = Session("ShipCity")
objRec("ShipProvince") = Session("ShipProvince")
objRec("ShipCountry") = Session("ShipCountry")
objRec("ShipPostalCode") = Session("ShipPostalCode")
objRec("ShipEmail") = Session("ShipEmail")
objRec("ShipTelephone") = Session("ShipTelephone")
objRec("ShipFax") = Session("ShipFax")

objRec("BillPrefix") = Session("BillPrefix")
objRec("BillFirstName") = Session("BillFirstName")
objRec("BillLastName") = Session("BillLastName")
objRec("BillCompany") = Session("BillCompany")
objRec("BillStreetAddress") = Session("BillStreetAddress")
objRec("BillStreetAddress2") = Session("BillStreetAddress2")
objRec("BillCity") = Session("BillCity")
objRec("BillProvince") = Session("BillProvince")
objRec("BillCountry") = Session("BillCountry")
objRec("BillPostalCode") = Session("BillPostalCode")
objRec("BillEmail") = Session("BillEmail")
objRec("BillTelephone") = Session("BillTelephone")
objRec("BillFax") = Session("BillFax")

objRec("CardType") = Session("CardType")

' Successfully billed the credit card, so hide the
' number now.
objRec("PaidDate") = Session("PaidDate")
objRec("CardNumber") = HideCreditCardNumber(Session("CardNumber"))
Session("CardNumber") = HideCreditCardNumber(Session("CardNumber"))
objRec("CardExpiry") = Session("CardExpiry")
objRec("CardName") = Session("CardName")
objRec("Items") = ItemStr
objRec("Currency") = Session("Currency")
objRec("CurrencyAbrev") = Session("CurrencyAbrev")
objRec("CurrencyExchange") = GetExchangeRate(Session("CurrencyAbrev"))
objRec("SubTotalCost") = Session("SubTotalCost")
objRec("ShippingCost") = Session("ShippingCost")
objRec("HSTCost") = Session("HSTCost")
objRec("GSTCost") = Session("GSTCost")
objRec("PSTCost") = Session("PSTCost")
objRec("TotalCost") = Session("TotalCost")
objRec("OriginalTotalCost") = Session("TotalCost")
objRec("ShippingRequired") = Session("ShippingRequired")
objRec("ClientNotes") = Session("ClientNotes")
objRec("PurchaseOrder") = Session("PurchaseOrder")
objRec("CardAuthorizationNo") = Session("CardAuthorizationNo")
If Session("PayQuoteMode") = "Yes" Then
    If Left(Session("CurrentTransaction"), 1) = "Q" Then
        objRec("OrderDate") = Date
    Else
        objRec("OrderDate") = Session("OrderDate")
        objRec("ProcessedDate") = Session("ProcessedDate")
    End If
    ObjRec("ContactID") = Session("ContactID")
    ObjRec("OfficeID") = Session("OfficeID")
    ObjRec("CompanyID") = Session("CompanyID")
Else
    objRec("OrderDate") = Date
End If
objRec("TransactionNumber") = NewTransactionNumber
'objRec("Quote") = ""

objRec("Source") = Session("SourceOption")

' auto set sales person base on source (for scott)
If Session("SourceOption") = 15 Then
    objRec("SalesPerson") = "Scott Jeffers"
End If


    ' ip address
    ipaddress = Request.ServerVariables("REMOTE_ADDR")
    If Len(ipaddress) > 50 Then ipaddress = Left(ipaddress, 50)
    
    objRec("IPAddress") = ipaddress

    objRec("ConsentToEmail") = Session("Consent")
' for conference interests
If Not IsNull(Session("ConferenceInterests")) And Len(Session("ConferenceInterests")) > 2 Then
    objRec("ConfirmationNotes") = Session("ConferenceInterests")
End If
' for conference permission (starting 2019's conference)
If Not IsNull(Session("ConferencePermission")) And Len(Session("ConferencePermission")) > 2 Then
    objRec("PCSWMMReference") = Session("ConferencePermission")
End If

If Not IsNull(Session("OurNotes")) And Len(Session("OurNotes")) > 2 Then
    objRec("OurNotes") = Session("OurNotes")
End If

	objRec.Update

	Response.LCID = 4105
	set JSONarr = New JSONArray

	Call JSONarr.LoadRecordSet(objRec)
	objRec.Close
	Set objRec = Nothing

	JSONarr(0).remove("TransactionNumber")
	JSONarr(0).Add "Company_ID", Session("CompanyID")
	JSONarr(0).Add "Name", NewTransactionNumber

	dim json
	json = JSONarr(0).Serialize()

	'Response.Write json
	'Response.End

	Set HttpReq = Server.CreateObject("MSXML2.ServerXMLHTTP")
	HttpReq.open "POST", apiUri & "InsertTransaction", False
	HttpReq.setRequestHeader "Content-Type", "application/json"
	HttpReq.setRequestHeader "Accept", "application/json"
	HttpReq.setRequestHeader "ApiKey", ApiKey
	HttpReq.send json

        ' add workshop interests
        If Session("WorkshopInterests") = "Yes" And Len(workshopItem) > 0 Then
            Set objRecSurvey = Server.CreateObject ("ADODB.Recordset")
            objRecSurvey.Open "PreWorkshopSurvey", surveyPath, 3, 3, 2

            objRecSurvey.AddNew
            objRecSurvey("Workshop") = workshopItem
            objRecSurvey("TransactionNumber") = NewTransactionNumber
            objRecSurvey("ModelingSoftwareExperience") = Session("WExperience1")
            objRecSurvey("PCSWMMExpercience") = Session("WExperience2")
            objRecSurvey("WorkshopTopic1") = Session("WInterest1")
            objRecSurvey("WorkshopTopic2") = Session("WInterest2")
            objRecSurvey("WorkshopTopic3") = Session("WInterest3")
            objRecSurvey.Update

            objRecSurvey.Close
            Set objRecSurvey = Nothing
        End If

Session("CurrentTransaction")= NewTransactionNumber

If payOrder Then
	EmailTitlePreStr = "Payment submitted for "
Else
	EmailTitlePreStr = "New "
End If
Session("PayQuoteMode") = ""

If isTrial Then
    EmailTitleStr = "Trial license request from " & Session("ClientFirstName") & " " &  Session("ClientLastName") & ", " & Session("ClientCompany") & ", " & Session("ClientCountry")
ElseIf isWebinar Then
    EmailTitleStr = "Webinar attendance from " & Session("ClientFirstName") & " " &  Session("ClientLastName") & ", " & Session("ClientCompany") & ", " & Session("ClientCountry")
Else
    EmailTitleStr = EmailTitlePreStr  & "order " & Session("CurrencyAbrev") & FormatCurrency(Session("TotalCost")) & " from " & Session("ClientFirstName") & " " &  Session("ClientLastName") & ", " & Session("ClientCompany") & ", " & Session("ClientCountry")
End If

'If workshopItem <> "" Then EmailTitleStr = Left(workshopItem, 4) & " - " & EmailTitleStr
If workshopItem <> "" Then EmailTitleStr = workshopItem & " - " & EmailTitleStr

Session("isTrial") = isTrial
EmailStr = EmailStr & vbCrLf
'send notice to Bill & Rob by email
											    If Session("ShippingRequired") = "Yes" Then
												EmailStr = EmailStr & vbCrLf & "Shipping Estimate: " & Session("ShippingEstimateMethod") & vbCrLf
											    End If
												EmailStr = EmailStr & "Payment method: " & Session("CardType") & vbCrLf & _
												"Confirmation #: " & Session("CardAuthorizationNo") & vbCrLf & vbCrLf & _
												"https://secure.chiwater.com/chi/orders/order.asp?Transaction=" & NewTransactionNumber

    If hasMultipleItem Then EmailStr = EmailStr & vbCrLf & vbCrLf & "This order has multiple items."

    ' indicate existing user
    numTransaction = 0
        set connTrans=Server.CreateObject("ADODB.Connection")
		connTrans.Provider="Microsoft.Jet.OLEDB.4.0"
		connTrans.Open transcationPath
	    strSQL = "Select Count(TransactionNumber) From [ProcessedTransactions] Where TotalCost > 0 and ClientEmail = ? "

        set objCommand = Server.CreateObject("ADODB.Command")
        objCommand.ActiveConnection = connTrans
        objCommand.CommandText = strSQL

        set objDbParam = objCommand.CreateParameter("@email", 200, 1, 50)
        objCommand.Parameters.Append objDbParam
        objCommand.Parameters("@email") = Session("ClientEmail")
        set objDbParam = Nothing
       
        set objRS = objCommand.Execute()

        If Not objRS.BOF And Not objRS.EOF Then
            numTransaction = objRS(0) + 1
            EmailStr = EmailStr & vbCrLf & vbCrLf & "This is their " & numTransaction & "  non-zero transaction(s) with CHI."
        End If 


	    objRS.Close
	    Set objRS = Nothing
        set objCommand = Nothing
	    connTrans.Close

    ' indicate existing user ends

emailTo = "ian@chiwater.com" '"rob@chiwater.com; info@chiwater.com;"
'If Session("TotalCost") > 0 Then emailTo = emailTo & " bill@chiwater.com;"
' indicate existing client with * in the email title
If numTransaction = 0 Then EmailTitleStr = "*" & EmailTitleStr

    ' to test ''"store@chiwater.com", _
	SendMail emailTo, _
			"ian@chiwater.com", _		
			EmailTitleStr, _
			EmailStr

	'SendMail "sharon@chiwater.com", _
	'		"store@chiwater.com", _
	'		"test orders for sharon - " & EmailTitleStr, _
	'		EmailStr

	Response.Redirect "orderconfirmation.asp"

' Return the transaction number from either the "CurrentTransaction" session variable, or
' check the database for the next transaction number.
Public Function GetTransactionNumber()
	If Session("Quote") = "Yes" Or Session("PayQuoteMode") = "Yes" Then
		' check pay order or quote
		If Not Left(Session("CurrentTransaction"), 1) = "Q" Then
					GetTransactionNumber = Session("CurrentTransaction")
					payOrder = True
					Exit Function
		End If
	End If

	Set objRec = Server.CreateObject ("ADODB.Recordset")
		set connTrans=Server.CreateObject("ADODB.Connection")
		connTrans.Provider="Microsoft.Jet.OLEDB.4.0"
		connTrans.Open transcationPath
	strSQL = "Select [TransactionNumber] From [Transaction] Where [TransactionNumber] Not Like 'Q%' And [TransactionNumber] Not Like 'G%' Order By Key Desc"
	'objRec.Open strSQL, transcationPath, 3, 3, 0
		objRec.Open strSQL, connTrans
	objRec.MoveFirst

	GetTransactionNumber = CLng(objRec("TransactionNumber")) + 1

	objRec.Close
	Set objRec = Nothing
	connTrans.Close
End Function

%>







