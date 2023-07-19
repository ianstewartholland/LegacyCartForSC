
            
    <%
    
	noteStr = "<tr><td valign='top' bgcolor='#FFFF99' colspan='2'>"
    

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

	
	noteStr = noteStr & "<p><font face='Arial' size='2'><b>Our notes</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font> "
	
            noteStr = noteStr & "<font face='Arial' size='1'><a href='/chi/notes/newnote.asp?Type=" & NoteType & "&ID=" & NoteRecordID & "'>New note</a></font>"
            If objRecNotes.EOF then
                noteStr = noteStr & "<font face='Arial' size='2'><br>(no notes, would you like to <a href='/chi/notes/newnote.asp?Type=" & NoteType & "&ID=" & NoteRecordID & "'>create a note</a>?)</font>"
            Else
                noteStr = noteStr & "<br>"
                noteStr = noteStr & "<table>"
                Do While Not objRecNotes.EOF
                
                    noteStr = noteStr & "<tr>"
                        
                        noteStr = noteStr & "<td valign='top' style='width:80px; position:relative; vertical-align: baseline;'><font face='Arial' size='1'>"
                        
                        If IsNull(objRecNotes("Date")) then
                            noteStr = noteStr & "<b>(Earlier Note)</b></font></td>"
                        Else
                            noteStr = noteStr & objRecNotes("Date")
                            noteStr = noteStr & "<br>" & objRecNotes("Author") & "</font>"
                            noteStr = noteStr & "</td>"
                        End If
                        
                        noteStr = noteStr & "<td valign='top' style='margin-top:0; padding-top: 0px; position:relative; vertical-align: baseline;'><font face='Arial' size='2'>" & InsertTags(objRecNotes("Content")) & "</font> "
                        noteStr = noteStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
                        noteStr = noteStr & "<font face='Arial' size='1'><a href='/chi/notes/editnote.asp?Type=" & NoteType & "&ID=" & NoteRecordID & "&ThisNoteID=" & objRecNotes("NotesID") & "'>Edit</a> </font>"
                        noteStr = noteStr & "<font face='Arial' size='1'><a href='/chi/notes/deletenote.asp?Type=" & NoteType & "&ID=" & NoteRecordID & "&ThisNoteID=" & objRecNotes("NotesID") & "'>Delete</a></font>"
                        noteStr = noteStr & "</td>"
                        
                    noteStr = noteStr & "</tr>"  
    
                    objRecNotes.MoveNext
                Loop
                noteStr = noteStr & "</table>"
                
            End If
           
	    objRecNotes.Close
  	    Set objRecNotes = Nothing
	    
	noteStr = noteStr & "</p></td></tr>"

	%>