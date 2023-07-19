<!--#include virtual ="/chi/includes/AddItem.asp"-->
<%
'   
' use this page to add item directly into the cart and jump to client information page
' item needs to be pre defined in the backend
'
    
    Session("QuoteQuery") = "Workshops"
    
    'If Len(Request.QueryString("Item")) = 4 And Left(Request.QueryString("Item"), 1) = "W" And Request.QueryString("Item") <> "W842" Then
    If Len(Request.QueryString("Item")) = 4 And Request.QueryString("Item") = "W953" Then
        AddItem Request.QueryString("Item"), 1

        'edit item description
        Session("CartItemDescription(1)") = Replace(Session("CartItemDescription(1)"), "(must be registered for EGU conference, please <a href='mailto:info@chiwater.com'>email</a> proof of registration)", "") 
        Session("CartItemDescription(1)") = Replace(Session("CartItemDescription(1)"), "(1 day)", "") & "(must be registered for EGU conference, please <a href='mailto:info@chiwater.com'>email</a> proof of registration)"
        Session("CartItemDescription(1)") = Replace(Session("CartItemDescription(1)"), "training", "<a target='_blank' href='http://meetingorganizer.copernicus.org/EGU2017/session/25610'>training</a>")
    	'add subtotal for the next page
        t = 1
        Session("SubTotalCost") = Session("CartItemCost(" & t & ")")
        Session("TotalCost") = Session("CartItemCost(" & t & ")")
    Else
        Response.Redirect "http://www.chiwater.com/"
    End If 
    


    Response.Redirect "cart.asp"

%>




