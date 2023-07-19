<!--#include virtual ="/chi/includes/AddItem.asp"-->
<%

	'check to process form
	If Session("PCSWMMNumLicenses") > 0 Then
	Else
		'Session("PCSWMMSoftwareVersion") = Request.Form("SoftwareVersion")
		'Session("PCSWMMLicenseType") = Request.Form("LicenseType")
		'Session("PCSWMMNumLicenses") = Request.Form("Licenses")
		'Session("PCSWMMNumManuals") = Request.Form("Manuals")
		'Session("PCSWMMReference") = Request.Form("Reference")
		PCSWMMSoftwareVersion = Request.Form("SoftwareVersion")
		PCSWMMLicenseType = Request.Form("LicenseType")
		Session("PCSWMMNumLicenses") = Request.Form("Licenses")
		PCSWMMNumManuals = Request.Form("Manuals")
	End If
	
	'check to ensure the Catalog is loaded
	'If Session("CatalogLoaded") <> "Yes" Then 
	'  	Session("NextPage") = "pcswmmnetaddtocartcommit.asp"
	'	Response.Redirect("Catalog.asp")
	'End If
	
	'remove any existing software items from cart
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
	Select Case PCSWMMSoftwareVersion
		Case "PCSWMM 2011 Standard"
			Select Case PCSWMMLicenseType
				Case "Single-User License"
					'add PCSWMM 2011 Standard single user licenses
					AddItem "S201", Session("PCSWMMNumLicenses")
				Case "Enterprise License"
					'add PCSWMM 2011 Standard Enterprise license
					AddItem "S205", 1
				Case "Educational Multi-User (Classroom) License"
					'not a valid option
			End Select
		Case "PCSWMM 2011 Professional"
			Select Case PCSWMMLicenseType
				Case "Single-User License"
					'add PCSWMM 2011 Professional single user licenses
					AddItem "S202", Session("PCSWMMNumLicenses")
				Case "Enterprise License"
					'add PCSWMM 2011 Professional Enterprise license
					AddItem "S206", 1
				Case "Educational Multi-User (Classroom) License"
					'not a valid option
			End Select
		Case "PCSWMM 2010 Real-Time"
			Select Case PCSWMMLicenseType
				Case "Single-User License"
					'add PCSWMM 2011 Real-Time single user licenses
					AddItem "S203", Session("PCSWMMNumLicenses")
				Case "Enterprise License"
					'not a valid option
				Case "Educational Multi-User (Classroom) License"
					'not a valid option
			End Select
		Case "PCSWMM 2011 Educational"
			Select Case PCSWMMLicenseType
				Case "Single-User License"
					'add PCSWMM 2011 Educational single-user licenses
					AddItem "S210", Session("PCSWMMNumLicenses")
				Case "Enterprise License"
					'not a valid option
				Case "Educational Multi-User (Classroom) License"
					'add PCSWMM 2011 Educational multi-user license
					AddItem "S211", 1
			End Select
	End Select
	
	'add discount items to cart
	'If Session("PCSWMMSoftwareVersion") <> "PCSWMM 2010 Real-Time" And Session("PCSWMMLicenseType") = "Single-User License" Then
	'	Select Case Session("PCSWMMCurrentSoftwareVersion")
	'		Case "PCSWMM 2006"
	'			Select Case Session("PCSWMMSoftwareVersion")
	'				Case "PCSWMM 2011 Standard"
	'					'add PCSWMM 2011 Standard upgrade discount (20%)
	'					AddItem "S902", Session("PCSWMMCurrentNumLicenses")
	'				Case "PCSWMM 2011 Professional"
	'					'add PCSWMM 2011 Professional upgrade discount (20%)
	'					AddItem "S904", Session("PCSWMMCurrentNumLicenses")
	'			End Select
	'		Case "PCSWMM 2005 or earlier"
	'			Select Case Session("PCSWMMSoftwareVersion")
	'				Case "PCSWMM 2011 Standard"
	'					'add PCSWMM 2011 Standard upgrade discount (10%)
	'					AddItem "S901", Session("PCSWMMCurrentNumLicenses")
	'				Case "PCSWMM 2011 Professional"
	'					'add PCSWMM 2011 Professional upgrade discount (10%)
	'					AddItem "S903", Session("PCSWMMCurrentNumLicenses")
	'			End Select
	'	End Select
	'End If
	
	'add SWMM manuals to cart
	If PCSWMMNumManuals > 0 Then AddItem "R242", PCSWMMNumManuals
	
	'remove temporary session variables
	'Session("PCSWMMSoftwareVersion") = ""
	'Session("PCSWMMLicenseType") = ""
	'Session("PCSWMMNumLicenses") = 0
	'Session("PCSWMMCurrentSoftwareVersion") = ""
	'Session("PCSWMMCurrentLicenseType") = ""
	'Session("PCSWMMCurrentNumLicenses") = 0
	'Session("PCSWMMNumManuals") = 0
	'Session("PCSWMMReference") = ""
	
	'Sub AddItem(ItemStr, ItemNum)
	'	'find item
	'	For p = 1 to Session("NumCatalogItems")
	'		If Session("CatalogItem(" & p & ")") = ItemStr Then 
	'			t = t + 1
	'			Session("NumCartItems") = t
	'			Session("CartItem(" & t & ")") = ItemStr
	'			Session("CartItemDescription(" & t & ")") = Session("CatalogItemDescription(" & p & ")")
	'			Session("CartItemCost(" & t & ")") = Session("CatalogItemCost(" & p & ")")
	'			Session("CartItemQuantity(" & t & ")") = ItemNum
	'			Exit For
	'		End If
	'	Next
	'End Sub
	
	'set flags for recalculating
	'Session("SubTotalCalculated") = ""
	'Session("ShippingCalculated") = ""
	'Session("ShippingRequired") = ""
	'Session("TaxCalculated") = ""
	'Session("TotalCost") = 0
	
	Response.Redirect "cart.asp"

%>




