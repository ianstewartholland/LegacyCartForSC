<!--#include virtual ="/chi/includes/AddItem.asp"-->
<%

	'check to process form
	'If Session("PCSWMMNumLicenses") > 0 Then
	'Else
		PCSWMMNumLicenses = Request.Form("Licenses")
		PCSWMMSoftwareVersion = Request.Form("SoftwareVersion")
		'Session("PCSWMMNumLicenses") = Request.Form("Licenses")
		PCSWMMNumManuals = Request.Form("Manuals")
	'End If
    
    Session("QuoteQuery") = "PCSWMM"
	
	Do
		For t = 1 to Session("NumCartItems")
			If Left(Session("CartItem(" & t & ")"),1) = "S" Or Session("CartItem(" & t & ")") = "R242" Then Exit For
		Next
		If t > Session("NumCartItems") Then Exit Do
		'software item found
		Session("NumCartItems") = Session("NumCartItems") - 1
		For n = t to Session("NumCartItems")
			Session("CartItem(" & n & ")") = Session("CartItem(" & n + 1 & ")")
			Session("CartItemDescription(" & n & ")") = Session("CartItemDescription(" & n + 1 & ")")
			Session("CartItemCost(" & n & ")") = Session("CartItemCost(" & n + 1 & ")")
			Session("CartItemQuantity(" & n & ")") = Session("CartItemQuantity(" & n + 1 & ")")
		Next
		Session("CartItem(" & n & ")") = ""
		Session("CartItemDescription(" & n & ")") = ""
		Session("CartItemCost(" & n & ")") = 0
		Session("CartItemQuantity(" & n & ")") = 0
	Loop

	t = Session("NumCartItems")
	If t > 0 Then
	Else
		t = 0
	End If


	'add software items to cart
        If PCSWMMNumLicenses > 0 Then
            Select Case PCSWMMSoftwareVersion
                    Case "PCSWMM Professional"
                        AddItem "S220", PCSWMMNumLicenses                      
                        t = Session("NumCartItems")
                        If InStr(Session("CartItemDescription(" & t & ")"), "(") > 1 Then Session("CartItemDescription(" & t & ")") = Mid(Session("CartItemDescription(" & t & ")"), 1, InStr(Session("CartItemDescription(" & t & ")"), "(") - 1)
                        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & " (" & FormatDateStr() & ")"

                    Case "PCSWMM Professional 2D"
                        AddItem "S222", PCSWMMNumLicenses
                        t = Session("NumCartItems")
                        If InStr(Session("CartItemDescription(" & t & ")"), "(") > 1 Then Session("CartItemDescription(" & t & ")") = Mid(Session("CartItemDescription(" & t & ")"), 1, InStr(Session("CartItemDescription(" & t & ")"), "(") - 1)
                        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & " (" & FormatDateStr() & ")" 
                        
                    Case "PCSWMM Enterprise"
                        AddItem "S236", 1
                        t = Session("NumCartItems")
                        If InStr(Session("CartItemDescription(" & t & ")"), "(") > 1 Then Session("CartItemDescription(" & t & ")") = Mid(Session("CartItemDescription(" & t & ")"), 1, InStr(Session("CartItemDescription(" & t & ")"), "(") - 1)
                        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & " (" & FormatDateStr() & ")"
                        
                        AddItem "S237", PCSWMMNumLicenses
                        t = Session("NumCartItems")
                        If InStr(Session("CartItemDescription(" & t & ")"), "(") > 1 Then Session("CartItemDescription(" & t & ")") = Mid(Session("CartItemDescription(" & t & ")"), 1, InStr(Session("CartItemDescription(" & t & ")"), "(") - 1)
                        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & " (" & FormatDateStr() & ")" 
            End Select
        Else
            Response.Redirect "http://www.chiwater.com/Software/PCSWMM/OrderCalculator.asp"
        End If
	
	'add SWMM manuals to cart
	If PCSWMMNumManuals > 0 Then AddItem "R242", PCSWMMNumManuals
	
	Response.Redirect "cart.asp"
Function FormatDateStr()
    toDate = DateAdd("yyyy", 1, Date)
    toDate = DateAdd("d", -1, toDate)
    FormatDateStr = MonthName(Month(Date)) & " " & Day(Date) & ", " & Year(Date) & " to " & MonthName(Month(toDate)) & " " & Day(toDate) & ", " & Year(toDate)
End Function
%>




