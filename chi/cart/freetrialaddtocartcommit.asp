<!--#include virtual ="/chi/includes/AddItem.asp"-->
<%
    
'    If Session("CatalogLoaded") <> "Yes" Then 
'  	Session("NextPage") = "freetrialaddtocartcommit.asp"
'	Response.Redirect("Catalog.asp")
'    End If
    
    Session("QuoteQuery") = "PCSWMM"
    
    t = Session("NumCartItems")
    If t > 0 Then
    Else
        t = 0
    End If
    
    AddItem "S212", 1
    
    'Sub AddItem(ItemStr, ItemNum)
    '        'find item
    '        For p = 1 to Session("NumCatalogItems")
    '                If Session("CatalogItem(" & p & ")") = ItemStr Then 
    '                        t = t + 1
    '                        Session("NumCartItems") = t
    '                        Session("CartItem(" & t & ")") = ItemStr
    '                        Session("CartItemDescription(" & t & ")") = Session("CatalogItemDescription(" & p & ")")
    '                        Session("CartItemCost(" & t & ")") = Session("CatalogItemCost(" & p & ")")
    '                        Session("CartItemQuantity(" & t & ")") = ItemNum
    '                        Exit For
    '                End If
    '        Next
    'End Sub
    
    'set flags for recalculating
    'Session("SubTotalCalculated") = ""
    'Session("ShippingCalculated") = ""
    'Session("ShippingRequired") = ""
    'Session("TaxCalculated") = ""
    'Session("TotalCost") = 0
    
    Response.Redirect "clientaddress.asp"

%>




