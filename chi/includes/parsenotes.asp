<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<%
    ' this is find source page type and id 
    NoteType = Trim(Request.QueryString("Type"))
    NoteRecordID = Trim(Request.QueryString("ID"))
    ThisNoteID = Trim(Request.QueryString("ThisNoteID"))

    ' find out database path and the next page to redirect
    Select Case LCase(NoteType)
        Case "conference"
            NoteDBPath = conferencePath
            nextpage = "/chi/conferences/" & NoteType & ".asp?" & NoteType & "=" & NoteRecordID
        Case "workshop"
            NoteDBPath = workshopPath
            'nextpage = "/chi/workshops/" & NoteType & ".asp?" & NoteType & "=" & NoteRecordID
	    '/chi/orders/crmlinks.asp?Type=workshop&WorkshopNumber=
	    'nextpage = "/chi/orders/crmlinks.asp?Type=workshop&WorkshopNumber=" & NoteRecordID
	    '/crm/workshop.aspx?Workshop=
	    nextpage = "/chi/crm/workshops/workshop.aspx?Workshop=" & NoteRecordID
	    ' see if there is a select para
	    If Len(Request.QueryString("select")) > 0 Then
		nextpage = nextpage & "&select=" & Request.QueryString("Select")
	    End If
        Case "transaction"
            NoteDBPath = transcationPath
            nextpage = "/chi/orders/order.asp?" & NoteType & "=" & NoteRecordID
	Case "grant"
            NoteDBPath = transcationPath
            nextpage = "/chi/crm/grant/grant.aspx?transaction=" & NoteRecordID
        Case "license"
            NoteDBPath = licensePath
            nextpage = "/chi/licenses/license.asp?" & NoteType & "=" & NoteRecordID
        Case "enterpriselicense"
            NoteDBPath = licensePath
            nextpage = "/chi/licenses/enterpriselicense.asp?" & NoteType & "=" & NoteRecordID
	Case "francelicense"
            NoteDBPath = licensePath
            nextpage = "/chi/france/license.asp?License=" & NoteRecordID
	Case "francecompanylicense"
            NoteDBPath = licensePath
            nextpage = "/chi/france/companylicense.asp?FranceLicense=" & NoteRecordID  
        Case "survey"
            NoteDBPath = surveyPath
            nextpage = "/chi/workshops/gateway.asp?Type=survey&ID=" & NoteRecordID & "&Year=" & Trim(Request.QueryString("Year")) & "&Workshop=" & Trim(Request.QueryString("Workshop"))
    End Select
%>