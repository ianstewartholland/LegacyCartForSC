<tr>
    <td valign="top" bgcolor="#FFFF99" colspan="2">
            
    <%
        NotesSQL = "SELECT NotesID, Date, RecordID, Content, Author"
        If NoteType = "francecompanylicense" Then
            NotesSQL = NotesSQL & " FROM FranceNotes WHERE RecordID LIKE '" & NoteRecordID & "' ORDER BY Date DESC"
        ElseIf NoteType = "enterpriselicense" Then
            NotesSQL = NotesSQL & " FROM EnterpriseNotes WHERE RecordID LIKE '" & NoteRecordID & "' ORDER BY Date DESC"
	Else
	    NotesSQL = NotesSQL & " FROM Notes WHERE RecordID LIKE '" & NoteRecordID & "' ORDER BY Date DESC"
        End If
        
        Set objRecNotes = Server.CreateObject ("ADODB.Recordset")
        strConnectNotes = NoteDBPath
        objRecNotes.Open NotesSQL, strConnectNotes, 3, 3, 0
    %>
        
        <p>
         
            <font face="Arial" size="2"><b>Our notes</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font> 
        <%
            Response.Write "<font face='Arial' size='1'><a href='/chi/notes/newnote.asp?Type=" & NoteType & "&ID=" & NoteRecordID & "'>New note</a></font>"
            If objRecNotes.EOF then
                response.write("<font face='Arial' size='2'><br>(no notes, would you like to <a href='/chi/notes/newnote.asp?Type=" & NoteType & "&ID=" & NoteRecordID & "'>create a note</a>?)</font>")
            Else
                response.write "<br>"
                response.write "<table>"
                Do While Not objRecNotes.EOF
                
                    response.write "<tr>"
                        
                        response.write "<td valign='top' style='width:80px; position:relative; vertical-align: baseline;'><font face='Arial' size='1'>"
                        
                        If IsNull(objRecNotes("Date")) then
                            response.write "<b>(Earlier Note)</b></font></td>"
                        Else
                            response.write objRecNotes("Date")
                            response.write "<br>" & objRecNotes("Author") & "</font>"
                            response.write "</td>"
                        End If
                        
                        response.write "<td valign='top' style='margin-top:0; padding-top: 0px; position:relative; vertical-align: baseline;'><font face='Arial' size='2'>" & InsertTags(objRecNotes("Content")) & "</font> "
                        response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
                        response.write "<font face='Arial' size='1'><a href='/chi/notes/editnote.asp?Type=" & NoteType & "&ID=" & NoteRecordID & "&ThisNoteID=" & objRecNotes("NotesID") & "'>Edit</a> </font>"
                        response.write "<font face='Arial' size='1'><a href='/chi/notes/deletenote.asp?Type=" & NoteType & "&ID=" & NoteRecordID & "&ThisNoteID=" & objRecNotes("NotesID") & "'>Delete</a></font>"
                        response.write "</td>"
                        
                    response.write "</tr>"  
    
                    objRecNotes.MoveNext
                Loop
                response.write "</table>"
                
            End If
            objRecNotes.Close
	    Set objRecNotes = Nothing
        %>
            
        </p>  
    </td>
</tr>