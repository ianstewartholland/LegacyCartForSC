<!--#include virtual ="/chi/includes/pscheck.inc"-->
<table border="0" width="640" class = 'hiddenheader' >
        <tr>
            <td align="right">
                <font size="2" face="Arial">
                <%
                Select Case Session("PasswordLevel")
                    Case "Owner", "Admin"
                    %>
			<a href="/chi/crm/home.aspx">CRM</a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        <a href="/chi/orders/currentorders.asp">Orders</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="/chi/crm/grant/currentgrants.aspx">Grants</a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        <a href="/chi/licenses/currententerpriselicenses.asp">Licenses</a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        <a href="/chi/workshops/currentworkshops.asp">Workshops</a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        <a href="/chi/conferences/currentconferences.asp">Conferences</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<% If Session("PasswordLevel") = "Admin" Then %>
			    <a href="/chi/protected/hours/week.aspx">Hours</a><br>
			<% Else %>
			    <a href="/chi/protected/hours/summary.aspx">Hours</a><br>
			<% End If %>
			<a href="/chi/crm/test/editwebpage.aspx">Website</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="/chi/crm/security/security.aspx">Security</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="/chi/crm/files/index.aspx">Files</a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        <a href="/chi/orders/clickthroughs-summary.asp">Clickthroughs</a>
                    <%
		    Case "TechSupport"
		    %>
			<a href="/chi/crm/website/editwebpage.aspx">Website</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="/chi/crm/home.aspx">CRM</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="/chi/licenses/currententerpriselicenses.asp">Licenses</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="/chi/workshops/currentworkshops.asp">Workshops</a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        <a href="/chi/conferences/currentconferences.asp">Conferences</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="/chi/crm/files/index.aspx">Files</a>&nbsp;&nbsp;|&nbsp;&nbsp;
			<a href="/chi/protected/hours/week.asp">Hours</a>
		    <%
		    Case "Bookkeeper"
                ' below is set to be maria
                token = ""
                'If Session("CurrentEmployee") = "Maria Chovaz" Then token = "bWNob3Zhel4="
                'If Session("CurrentEmployee") = "Lindsay Kolbuc" Then token = "bGtvbGJ1Y14="
                'If Session("CurrentEmployee") = "Alex Thomas" Then token = "YXRob21hc14="
                'If Session("CurrentEmployee") = "Andrea Patteson" Then token = "YXBhdHRlc29uXg=="
                If Session("CurrentEmployee") = "Barbara Kerson" Then token = "YmtlcnNvbl4="
                If Session("CurrentEmployee") = "Joanna Mundy" Then token = "am11bmR5Xg=="
                '<a href="/chi/analysis/financials.asp">Analysis</a>&nbsp;&nbsp;|&nbsp;&nbsp;
                '<a href="/chi/crm/home.aspx">CRM</a>&nbsp;&nbsp;|&nbsp;&nbsp;
		    %>
                        <a href="/chi/orders/currentorders.asp">Orders</a>&nbsp;&nbsp;|&nbsp;&nbsp;
                        
                        
			<a href="/chi/protected/hours/summary.aspx?Current=<%=token%>">Hours</a>
		    <%
                End Select
                %>
                <%
                If Session("PasswordLevel") = "Owner" Then
                    %>
                        &nbsp;&nbsp;|&nbsp;&nbsp;
                        <a href="/chi/analysis/financials.asp">Analysis</a>
                    <%
                End If
                %>
                <br></font>
            </td>
        </tr>
</table>

