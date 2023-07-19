
<!--#include virtual ="/chi/includes/AddItem.asp"-->
<%
	'dim Item

	'get item
	'Item = Request.QueryString("Item")
	If Len(Request.QueryString("Item")) > 1 Then
	    'Session("CurrentItem") = Item
	    AddItem Request.QueryString("Item"), 1
	    
	    Response.Redirect "cart.asp"
	Else
	    Response.Redirect "http://www.chiwater.com/Publications/Books/index.asp"
	End If


	
'	'check for expired page
'	If Session("CurrentItem") = "" Then Response.Redirect "cart.asp"
'
'  	'store this page
'  	Session("NextPage") = "addtocart.asp"
'	
'	'check to ensure the Catalog is loaded
'	if Session("CatalogLoaded") <> "Yes" then Response.Redirect("Catalog.asp")
'  	
'	'set up the last page (if the user wants to cancel)
'	Select Case Left(Session("CurrentItem"),1)
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
'
'	'get the item index
'	For t = 1 to Session("NumCatalogItems")
'		If lcase(Session("CurrentItem")) = lcase(Session("CatalogItem(" & t & ")")) Then Exit For
'	Next
		
%>
