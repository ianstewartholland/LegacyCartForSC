<%
    '==============================
    ' Quote control
    '==============================

    NumQuotes = 10
    QuoteQuery = Session("QuoteQuery")

    
    dim objConn
    dim objRec
    dim strConnect
    dim strQuoteIndex
    dim strQuoteList
    dim lenQuoteList
    dim q
    
    strConnect ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WWWUsers\LocalUser\chi\computationalhydraulics_com\databases\testimonials.mdb;"
	    set conn=Server.CreateObject("ADODB.Connection")
	    conn.Provider="Microsoft.Jet.OLEDB.4.0"
	    conn.Open strConnect
    'connect to database
    Set objRec = Server.CreateObject ("ADODB.Recordset")	
    
    strQuoteIndex = QuoteQuery & "QuoteIndex"
    strQuoteList = QuoteQuery & "QuoteList"
    
    If NumQuotes > 0 Then

        If IsEmpty(Session(strQuoteIndex)) Then
        
            ' Initialize ASP RND() function
            Randomize()
            intRandomNumber = Int (1000*Rnd)+1
            
            objRec.Open "SELECT ID FROM [" & QuoteQuery & "] ORDER BY RND(" & -1 * (intRandomNumber) & "*ID)", conn ' strConnect, 3, 3, 1
            
            Session(strQuoteList) = objRec.GetRows()
            Session(strQuoteIndex) = 0
            objRec.Close
            
        End If
    %>
    <ul>
        <li class="leftbartop"></li>
        <li class="leftbarcontent">
            <br />
	    <%
            lenQuoteList = UBound(Session(strQuoteList),2)
            
            For q = 0 to 3
                Session(strQuoteIndex) = (Session(strQuoteIndex) + 1) Mod (lenQuoteList - 1)
                objRec.Open "SELECT * FROM [" & QuoteQuery & "] WHERE ID=" & CInt(Session(strQuoteList)(0, Session(strQuoteIndex))), conn ' strConnect, 3, 3, 1    
                
                QuoteName = Trim(objRec("FirstName") & " " & objRec("LastName"))
                QuoteInitials = ObjRec("Initials")
                QuoteCompany = ObjRec("Company")
                QuoteCity = ObjRec("City")
                QuoteProvince = ObjRec("Province")
                QuoteCountry = ObjRec("Country")
                QuoteText = ObjRec("Quote")
            
                Response.Write "<p><i>" & QuoteText & "</i></p>"
                'If QuoteInitials <> "" Then
                    Response.Write "<p align=right>- " '& QuoteInitials & ", "
                'End If
                If QuoteCity <> "" Then
                    Response.Write QuoteCity  & ", "
                End If
                If QuoteProvince <> "" Then
                    Response.Write QuoteProvince & ", "
                End If
                If QuoteCountry <> "" Then
                    Response.Write QuoteCountry & "</p>"
                End If
                Response.Write "<p>&nbsp;</p>"                
                objRec.Close
            Next
            
            Set objRec = Nothing
	    conn.Close
	    
            %>
            <br />
        </li>
        <li class="leftbarbottom"></li>
    </ul>

	<%
    
    End If
    
    %>
				

