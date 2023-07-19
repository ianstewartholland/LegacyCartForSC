<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<!--#include virtual ="/chi/includes/AddItem.asp"-->
<!--#include virtual ="/chi/includes/taxcalculator.asp"-->
<%
    ' under attack if figure out url parameters - update with parameters if you can
    license = Request.QueryString("l")
    email = Request.QueryString("p")
    contactID = Request.QueryString("d")
    item = Request.QueryString("i")
    refund = Request.QueryString("r")
    quan = 0
    If IsNumeric(Request.QueryString("q")) Then
        quan = Request.QueryString("q")
    End If

    ' need to update enterprise item here if there is a change
    EnterpriseItem = "S236"
    BaseItem = "S237"
    refundItemStr = "Refund on license " & license
    
    ' validate
    If IsNull(license) Or IsNull(email) Or IsNull(contactID) Or IsNull(item) Then
        Response.Redirect "http://www.chiwater.com/"
    End If
    
    If Len(license) <> 6 Or Not IsNumeric(license) Or Not (InStr(email,"@") > 0 And InStrRev(email,".") > InStr(email,"@")) Or Len(contactID) > 5 Or Not IsNumeric(contactID) Then
        Response.Redirect "http://www.chiwater.com/"
    End If

    'test
    'license = "810657"
    'email = "rob@chiwater.com"
    'contactID = "1343"
    'item = "S236"

    ' clear items
    If Not IsNull(Session("NumCartItems")) And Session("NumCartItems") > 0 Then
        For i = 1 To Session("NumCartItems")
            Session("CartItem(" & t & ")") = ""
		    Session("CartItemDescription(" & t & ")") = ""
		    Session("CartItemCost(" & t & ")") = 0
		    Session("CartItemQuantity(" & t & ")") = 0
        Next
        
        Session("NumCartItems") = 0
    End If

    Session("QuoteQuery") = "PCSWMM"

    ' database 
    Set objRec = Server.CreateObject ("ADODB.Recordset")
    sql = "Select * From License Where License = '" & license & "' And Email = '" & email & "' And ContactID = " & contactID
	objRec.Open sql, licensePath, 3, 3, 1
    If Not objRec.EOF Then
        objRec.MoveFirst  
        Session("ClientPrefix") = ObjRec("Prefix")
	    Session("ClientFirstName") = ObjRec("FirstName")
	    Session("ClientLastName") = ObjRec("LastName")
	    Session("ClientEmail") = ObjRec("Email")
	    Session("ClientCompany") = ObjRec("Company")
	    Session("ClientStreetAddress") = ObjRec("StreetAddress")
	    Session("ClientStreetAddress2") = ObjRec("StreetAddress2")
	    Session("ClientCity") = ObjRec("City")
	    Session("ClientProvince") = ObjRec("Province")
	    Session("ClientPostalCode") = ObjRec("PostalCode")
	    Session("ClientCountry") = ObjRec("Country")
	    Session("ClientTelephone") = ObjRec("Telephone")
	    Session("ClientFax") = ObjRec("Fax")
	    Session("ClientWebSite") = ObjRec("WebSite")
        Session("ClientType") = objRec("Type")
   
        Session("ClientName") = Session("ClientFirstName") & " " & Session("ClientLastName")
        Session("BillName") = Session("ClientName")
        Session("ShipName") = Session("ClientName") 
        
        ' most likely already in the crm
        Session("ContactID") = objRec("ContactID")
        Session("OfficeID") = ObjRec("OfficeID") 
        Session("CompanyID") = ObjRec("CompanyID") 

        Session("ShipPrefix") = Session("ClientPrefix") 
        Session("ShipFirstName") = Session("ClientFirstName")
        Session("ShipLastName") = Session("ClientLastName")
        Session("ShipCompany") = Session("ClientCompany")
        Session("ShipStreetAddress") = Session("ClientStreetAddress")
        Session("ShipStreetAddress2") = Session("ClientStreetAddress2")
        Session("ShipCity") = Session("ClientCity")
        Session("ShipProvince") = Session("ClientProvince")
        Session("ShipCountry") = Session("ClientCountry")
        Session("ShipPostalCode") = Session("ClientPostalCode")
        Session("ShipEmail") = Session("ClientEmail")
        Session("ShipTelephone") = Session("ClientTelephone")
        Session("ShipFax") = Session("ClientFax")
    
        Session("BillPrefix") = Session("ClientPrefix") 
        Session("BillFirstName") = Session("ClientFirstName")
        Session("BillLastName") = Session("ClientLastName")
        Session("BillCompany") = Session("ClientCompany")
        Session("BillStreetAddress") = Session("ClientStreetAddress")
        Session("BillStreetAddress2") = Session("ClientStreetAddress2")
        Session("BillCity") = Session("ClientCity")
        Session("BillProvince") = Session("ClientProvince")
        Session("BillCountry") = Session("ClientCountry")
        Session("BillPostalCode") = Session("ClientPostalCode")
        Session("BillEmail") = Session("ClientEmail")
        Session("BillTelephone") = Session("ClientTelephone")
        Session("BillFax") = Session("ClientFax")
    
        Session("ShippingCost") = 0

        ' put the item together
        AddItem item, 1
        If item = EnterpriseItem And quan > 0 Then
            AddItem BaseItem, quan
        End If

    ' see if there is a refund item
    If Not IsNull(refund) And Len(refund) > 5 And Left(refund, 1) = "S" And InStr(refund, "R") > 0 And InStr(refund, "|") > 0 Then
        ' do not refund if downgrading
        If item <> "S220" Then

            refundItemStr = refundItemStr & " (" & Mid(refund, InStr(refund, "|") + 1) & ")"

            t = Session("NumCartItems") + 1
            Session("NumCartItems") = t
            Session("CartItem(" & t & ")") = Left(refund, 5)
		    Session("CartItemDescription(" & t & ")") = refundItemStr
            
            refundItemStr = Replace(refund, Mid(refund, InStr(refund, "|")), "")
		    Session("CartItemCost(" & t & ")") = Mid(refundItemStr, 6) * -1
		    Session("CartItemQuantity(" & t & ")") = 1
        End If
    End If

        ' currency
        If Session("ShipCountry") = "Canada" Then
            Session("Currency") = "Canadian"
            Session("CurrencyAbrev") = "CAD"
        Else
            If Session("ShipCountry") = "USA" Then
                Session("Currency") = "USA"
            Else
                Session("Currency") = "International"
            End If 
            Session("CurrencyAbrev") = "US"
        End If
        
        ' item string
        ItemStr = ""
		For t = 1 to Session("NumCartItems")
		    ItemStr = ItemStr & Session("CartItem(" & t & ")") & "/" & Session("CartItemDescription(" & t & ")") & "/" & Session("CartItemCost(" & t & ")") & "/" & Session("CartItemQuantity(" & t & ")") & "/"
        Next

        If IsNull(Session("NoGST")) Then Session("NoGST") = ""
		If IsNull(Session("NoPST")) Then Session("NoGST") = ""
		If IsNull(Session("NoHST")) Then Session("NoGST") = ""

		Tax = TaxCalculator (Session("ShipProvince"), Session("CurrencyAbrev"), ItemStr, Session("ShippingCost"), Session("NoGST"), Session("NoPST"), Session("NoHST"))

		Session("GSTCost") = Tax(0)
		Session("PSTCost") = Tax(1)
		Session("HSTCost") = Tax(2)
		Session("TotalCost") = Tax(3)
		Session("SubTotalCost") = Tax(4)
    
	    If Session("SubTotalCost") = 0 Then
	        Session("TotalCost") = 0
	    End If
        
        ' other options
        Session("SourceOption") = 8
        
        Session("PayQuoteMode") = "No"
        Session("AddressClient") = True
        Session("ShowClient") = True
    End If

    objRec.Close
	Set objRec = Nothing

    Response.Redirect "https://secure.chiwater.com/chi/cart/checkout.asp"
%>