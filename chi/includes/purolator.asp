<!--#include file ="dbpaths.inc"-->
<!--#include file ="WriteLogEntry.asp"-->
<%
' GetPdsGetQuickShippingEstimate
' Send a SOAP request to the Purolator web service and receive a list
' of shipping estimates for the location specified.  Returns a
' list of shipping estimates, or Nothing is there is an error.
Public Function GetPdsQuickShippingEstimate( _
                        SenderPostalCode, _
                        ReceiverCity, _
                        ReceiverProvince, _
                        ReceiverCountry, _
                        ReceiverPostalCode, _
                        Packaging, _
                        Weight, _
                        WeightUnits)
                        
    Dim Development
    Dim serviceURL
    Dim Username
    Dim Password
    Dim BillingAcctNumber
    Dim XmlHTTP
    Dim myXML
    Dim SOAPRequest
    Dim success
    
    Development = 0 'development or production?
    If Development = 1 Then
        serviceURL = "https://devwebservices.purolator.com/PWS/V1/Estimating/EstimatingService.asmx"
        Username = "1dd76ddfdc01406f85b7851fec0e13ad" 'dev key
        Password = "7H}cxeBT" 'pass
        BillingAcctNumber = "9999999999"
    Else
        serviceURL = "https://webservices.purolator.com/PWS/V1/Estimating/EstimatingService.asmx"
        Username = "e45ec205771b400090a253dfc09a1ad3"
        Password = "v^uj7$FH"
        BillingAcctNumber = "4723597"
    End If
WriteLogEntry("in GetPdsQuickShippingEstimate...")    
    Set XmlHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    Set myXML =Server.CreateObject("MSXML.DOMDocument")
    myXML.Async=False
  
    XmlHTTP.Open "POST", serviceURL, False, Username, Password           
    XmlHTTP.setRequestHeader "Content-Type", "text/xml" 
    XmlHTTP.setRequestHeader "SOAPAction", "http://purolator.com/pws/service/v1/GetQuickEstimate"
    SOAPRequest = "<?xml version='1.0' encoding='utf-8'?>"
      
    SOAPRequest = SOAPRequest & "<soap-env:Envelope xmlns:ns1=""http://purolator.com/pws/datatypes/v1"" xmlns:soap-env=""http://schemas.xmlsoap.org/soap/envelope/"">"
    SOAPRequest = SOAPRequest & "<soap-env:Header>" & _
                                        "<ns1:RequestContext>" & _
                                                "<ns1:Version>1.0</ns1:Version>" & _
                                                "<ns1:Language>en</ns1:Language>" & _
                                                "<ns1:GroupID>xxx</ns1:GroupID>" & _
                                                "<ns1:RequestReference>Rating Example</ns1:RequestReference>" & _
                                        "</ns1:RequestContext>" & _
                                "</soap-env:Header>"
 
 
    SOAPRequest = SOAPRequest & "<soap-env:Body>" & _
                                        "<ns1:GetQuickEstimateRequest>" & _
                                                "<ns1:BillingAccountNumber>" & BillingAcctNumber & "</ns1:BillingAccountNumber>" & _
                                                "<ns1:SenderPostalCode>" & SenderPostalCode & "</ns1:SenderPostalCode>" & _
                                                "<ns1:ReceiverAddress>" & _
                                                        "<ns1:City>" & ReceiverCity & "</ns1:City>" & _
                                                        "<ns1:Province>" & ReceiverProvince & "</ns1:Province>" & _
                                                        "<ns1:Country>" & ReceiverCountry & "</ns1:Country>" & _
                                                        "<ns1:PostalCode>" & ReceiverPostalCode & "</ns1:PostalCode>" & _
                                                "</ns1:ReceiverAddress>" & _
                                                "<ns1:PackageType>" & Packaging & "</ns1:PackageType>" & _
                                                "<ns1:TotalWeight>" & _
                                                        "<ns1:Value>" & Weight & "</ns1:Value>" & _
                                                        "<ns1:WeightUnit>" & WeightUnits & "</ns1:WeightUnit>" & _
                                                "</ns1:TotalWeight>" & _
                                        "</ns1:GetQuickEstimateRequest>" & _
                                "</soap-env:Body>"
                                    
    SOAPRequest = SOAPRequest & "</soap-env:Envelope>"
    
    success = XmlHTTP.send(SOAPRequest)
    if success <> 0 then
        Set GetPdsQuickShippingEstimate = Nothing
        Exit Function
    end if
WriteLogEntry("success in GetPdsQuickShippingEstimate: " & success)   
    Set GetPdsQuickShippingEstimate = XmlHTTP.responseXML

    Set XmlHTTP = Nothing
End Function

' Parse the response XML from Purolator and get the
' price of ground shipping
Public Function GetPdsGroundShippingPrice(responseXML)
    Dim estimateXML
    Dim nodesFaults
    Dim nodesErrors
    Dim nodesTotalPrice
    Dim nodesServiceID

    Set estimateXML = Server.CreateObject("MSXML.DOMDocument")
    estimateXML.Async = False
    
    if Not estimateXML.load(responseXML) then
	Dim strErrText
        Dim xPE
	'Dim xPE = Server.CreateObject(MSXML.IXMLDOMParseError)
	' Obtain the ParseError object
	Set xPE = estimateXML.parseError
	With xPE
	    strErrText = "Your XML Document failed to load" & _
	    "due the following error." & vbCrLf & _
	    "Error #: " & .errorCode & ": " & xPE.reason & _
	    "Line #: " & .Line & vbCrLf & _  
	    "Line Position: " & .linepos & vbCrLf & _
	    "Position In File: " & .filepos & vbCrLf & _
	    "Source Text: " & .srcText & vbCrLf & _
	    "Document URL: " & .url
	End With
        response.write strErrText & "<br/>"
    end if

    Set nodesFaults = estimateXML.documentElement.selectNodes("//faultcode")
    if nodesFaults.Length >= 1 then
        GetPdsGroundShippingPrice = -1
        Exit Function
    end if
    
    Set nodesErrors = estimateXML.documentElement.selectNodes("//Errors")
    if nodesErrors.Length >= 1 and Trim(nodesErrors.Item(0).Text) <> "" then
        GetPdsGroundShippingPrice = -1
        Exit Function
    end if
    
    Set nodesTotalPrice = estimateXML.documentElement.selectNodes("//TotalPrice")
    Set nodesServiceID = estimateXML.documentElement.selectNodes("//ServiceID")

    ' Something large so that we can find the min price    
    GetPdsGroundShippingPrice = 1000000.00

    Dim i
    for i = 0 to nodesTotalPrice.Length - 1
        if InStr(nodesServiceID.Item(i).Text,"PurolatorGround") > 0  then
            Dim price
            price = CSng(nodesTotalPrice.Item(i).Text)
            if price < GetPdsGroundShippingPrice then
                GetPdsGroundShippingPrice = price
                Exit Function
            end if
        end if
    next
    
    ' Didn't find a ground shipping option
    GetPdsGroundShippingPrice = -1
    

End Function

' Parse the response XML from Purolator and get the
' lowest price of shipping
Public Function GetPdsLowestShippingPrice(responseXML)
    Dim estimateXML
    Dim nodesFaults
    Dim nodesErrors
    Dim nodesTotalPrice
    Dim nodesServiceID

    Set estimateXML = Server.CreateObject("MSXML.DOMDocument")
    estimateXML.Async = False
    
'    if Not estimateXML.load(responseXML) then
'	Dim strErrText
'        Dim xPE
'	'Dim xPE = Server.CreateObject(MSXML.IXMLDOMParseError)
'	' Obtain the ParseError object
'	Set xPE = estimateXML.parseError
'	With xPE
'	    strErrText = "Your XML Document failed to load" & _
'	    "due the following error." & vbCrLf & _
'	    "Error #: " & .errorCode & ": " & xPE.reason & _
'	    "Line #: " & .Line & vbCrLf & _  
'	    "Line Position: " & .linepos & vbCrLf & _
'	    "Position In File: " & .filepos & vbCrLf & _
'	    "Source Text: " & .srcText & vbCrLf & _
'	    "Document URL: " & .url
'	End With
'        response.write strErrText & "<br/>"
'    end if
    
    if Not estimateXML.load(responseXML) then
        GetPdsLowestShippingPrice = -1
        Exit Function
    end if
    
    ' mdj - Debugging (requires printxml.asp)
    ' DisplayNode estimateXML.childNodes
    Set nodesFaults = estimateXML.documentElement.selectNodes("//faultcode")
    if nodesFaults.Length >= 1 then
        GetPdsLowestShippingPrice = -1
        Exit Function
    end if
    
    Set nodesErrors = estimateXML.documentElement.selectNodes("//Errors")
    if nodesErrors.Length >= 1 and Trim(nodesErrors.Item(0).Text) <> "" then
        GetPdsLowestShippingPrice = -1
        Exit Function
    end if
    
    Set nodesTotalPrice = estimateXML.documentElement.selectNodes("//TotalPrice")
    Set nodesServiceID = estimateXML.documentElement.selectNodes("//ServiceID")

    ' Number large enough so that we can find the lowest shipping cost
    GetPdsLowestShippingPrice = 1000000.00

    Dim i
    for i = 0 to nodesTotalPrice.Length - 1
        Dim price
        if InStr(LCase(nodesServiceID.Item(i).Text), "ground") > 0 then
            price = CSng(nodesTotalPrice.Item(i).Text)
            if price < GetPdsLowestShippingPrice then
                GetPdsLowestShippingPrice = price
            end if
        end if
    next
    
    if GetPdsLowestShippingPrice = 1000000.00 then
        ' Didn't find a reasonable price
        GetPdsLowestShippingPrice = -1
    end if

End Function

' Parse the response XML from Purolator and get the
' lowest price of shipping via express.
Public Function PdsGetLowestExpressPrice(_
                                        SenderPostalCode, _
                                        ReceiverCity, _
                                        ReceiverProvince, _
                                        ReceiverCountry, _
                                        ReceiverPostalCode, _
                                        Packaging, _
                                        Weight, _
                                        WeightUnits)
    Dim respXML
    Dim estimateXML
    Dim nodesFaults
    Dim nodesErrors
    Dim nodesTotalPrice
    Dim nodesServiceID
    
    
    Set respXML = GetPdsQuickShippingEstimate(_
						SenderPostalCode, _
						ReceiverCity, _
						ReceiverProvince, _
						ReceiverCountry, _
						ReceiverPostalCode, _
						Packaging, _
						Weight, _
						WeightUnits)

	If respXML Is Nothing Then
WriteLogEntry("xml object: " & respXML)
		PdsGetLowestExpressPrice = -1
		Exit Function
	End If
	
    Set estimateXML = Server.CreateObject("MSXML.DOMDocument")
    estimateXML.Async = False
    
    if Not estimateXML.load(respXML) then
WriteLogEntry("did not load xml")  
        PdsGetLowestExpressPrice = -1
        Exit Function
    end if
 
    ' mdj - Debugging (requires printxml.asp)
    'DisplayNode estimateXML.childNodes
    Set nodesFaults = estimateXML.documentElement.selectNodes("//faultcode")
    if nodesFaults.Length >= 1 then
WriteLogEntry("nodesFaults.Length >= 1")  
        PdsGetLowestExpressPrice = -1
        Exit Function
    end if
    
    Set nodesErrors = estimateXML.documentElement.selectNodes("//Errors")
    if nodesErrors.Length >= 1 and Trim(nodesErrors.Item(0).Text) <> "" then
 WriteLogEntry("Trim(nodesErrors.Item(0).Text) is empty")     
        PdsGetLowestExpressPrice = -1
        Exit Function
    end if
    
    Set nodesTotalPrice = estimateXML.documentElement.selectNodes("//TotalPrice")
    Set nodesServiceID = estimateXML.documentElement.selectNodes("//ServiceID")

    ' Something large so that we can find the min price    
    PdsGetLowestExpressPrice = 1000000.00

    Dim i

    for i = 0 to nodesTotalPrice.Length - 1
        Dim price
        if InStr(LCase(nodesServiceID.Item(i).Text), "express") > 0 then
            price = CSng(nodesTotalPrice.Item(i).Text)
            if price < PdsGetLowestExpressPrice then
                PdsGetLowestExpressPrice = price
            end if
        end if
    next
     
    if PdsGetLowestExpressPrice = 1000000.00 then
  WriteLogEntry("no result")    
        ' Didn't find a reasonable price
        PdsGetLowestExpressPrice = -1
    end if

End Function

' Returns a Dictionary of "ShippingMethod" -> "Price"
' Returns True or False and the prices and methods dictionary in
' PriceList.
Public Function PdsGetShippingPrices(_
                        SenderPostalCode, _
                        ReceiverCity, _
                        ReceiverProvince, _
                        ReceiverCountry, _
                        ReceiverPostalCode, _
                        Packaging, _
                        Weight, _
                        WeightUnits, _
                        ByRef PriceList)
    Dim responseXML
	Dim estimateXML
    Dim nodesFaults
    Dim nodesErrors
    Dim nodesTotalPrice
    Dim nodesServiceID
	Dim priceInfo

    Set responseXML = GetPdsQuickShippingEstimate(_
						SenderPostalCode, _
						ReceiverCity, _
						ReceiverProvince, _
						ReceiverCountry, _
						ReceiverPostalCode, _
						Packaging, _
						Weight, _
						WeightUnits)

	If responseXML Is Nothing Then
		PdsGetShippingPrices = False
		Exit Function
	End If
    
    Set estimateXML = Server.CreateObject("MSXML.DOMDocument")
    estimateXML.Async = False
    
    if Not estimateXML.load(responseXML) then
        PdsGetShippingPrices = False
        Exit Function
    end if
    
    ' mdj - Debugging (requires printxml.asp)
    'DisplayNode estimateXML.childNodes
    Set nodesFaults = estimateXML.documentElement.selectNodes("//faultcode")
    if nodesFaults.Length >= 1 then
        PdsGetShippingPrices = False
        Exit Function
    end if
    
    Set nodesErrors = estimateXML.documentElement.selectNodes("//Errors")
    if nodesErrors.Length >= 1 and Trim(nodesErrors.Item(0).Text) <> "" then
        PdsGetShippingPrices = False
        Exit Function
    end if
    
    Set nodesTotalPrice = estimateXML.documentElement.selectNodes("//TotalPrice")
    Set nodesServiceID = estimateXML.documentElement.selectNodes("//ServiceID")

	Set PriceList = Server.CreateObject("Scripting.Dictionary")
	for i = 0 to nodesTotalPrice.Length - 1
		PriceList.Add nodesServiceID.Item(i).Text, nodesTotalPrice.Item(i).Text
 	next
    
    PdsGetShippingPrices = True
    
	
End Function

' Return the book's mass in kilograms (kg) or pounds (lb)
Function GetBookWeight(bookid, units)
    Dim strDBConnect
    Dim objConn
    Dim strSQL
    Dim objCommand
    Dim objDbParam
    Dim objRS

    ' Look up the book name in the Books biblio table and
    ' return the weight.
    'strDBConnect = "Driver={Microsoft Access Driver (*.mdb)};DBQ=C:\WWWUsers\LocalUser\chi\computationalhydraulics_com\databases\biblio_weights.mdb"
    strDBConnect = weightPath
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 15
    objConn.CommandTimeout = 30
    objConn.Open = strDBConnect
    
    strSQL = "SELECT [WEIGHT] FROM BOOKS WHERE BOOKID= ?;"
    
    set objCommand = Server.CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConn
    objCommand.CommandText = strSQL
    
    set objDbParam = objCommand.CreateParameter("@BookId", 200, 1, Len(bookid), bookid)
    objCommand.Parameters.Append objDbParam
    'objCommand.Parameters("@BookId") = bookid
    set objDbParam = Nothing

    set objRS = objCommand.Execute()
    
    if objRS.BOF or objRS.EOF then
        GetBookWeight = -1
    else
        objRS.MoveFirst
        GetBookWeight = objRS("Weight")
        if units = "kg" then
            GetBookWeight = GetBookWeight / 2.2
        end if
    end if
    
    set objCommand = Nothing
    objConn.Close
    set objConn = Nothing

End Function

Function ConvertCountryToCountryCode(Country)
    ' Look up the purolator two-letter country code for the
    ' country given by the customer.
    Dim strDBConnect
    Dim objConn
    Dim strSQL
    Dim objCommand
    Dim objDbParam
    Dim objRS    
    
    'strDBConnect = "Driver={Microsoft Access Driver (*.mdb)};DBQ=C:\WWWUsers\LocalUser\chi\computationalhydraulics_com\databases\ShippingCalculator2010.mdb"
    strDBConnect = shippingcalculatorPath
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 15
    objConn.CommandTimeout = 30
    objConn.Open = strDBConnect
    
    strSQL = "SELECT CODE FROM SHIPPINGCALCULATOR WHERE COUNTRY= ?;"
    
    set objCommand = Server.CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConn
    objCommand.CommandText = strSQL
    
    set objDbParam = objCommand.CreateParameter("@Country", 200, 1, Len(Country), Country)
    objCommand.Parameters.Append objDbParam
    'objCommand.Parameters("@BookId") = bookid
    set objDbParam = Nothing

    set objRS = objCommand.Execute()
    Set countryDict = Server.CreateObject("Scripting.Dictionary")
    
    if objRS.BOF or objRS.EOF then
        ConvertCountryToCountryCode = Null
    else
        'objRS.MoveFirst
        ConvertCountryToCountryCode = objRS("Code")
    end if
    
    set objCommand = Nothing
    objConn.Close
    set objConn = Nothing

End Function



%>
