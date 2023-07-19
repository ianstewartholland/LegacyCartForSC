<!--#include virtual ="/chi/includes/pscheck.inc"-->
<%
'
' check client related purchase
'
Function ClientCheck(tDate, cEmail, cPostal, comName)
    Dim clientCost
    Dim officeCost
    Dim companyCost
    clientCost = 0
    officeCost = 0
    companyCost = 0
    
    If InStr(comName, "'") Then comName = Replace(comName, "'", "''")
    ' dunno this ever happens
    If InStr(cEmail, "'") Then cEmail = Replace(cEmail, "'", "''")
    
    Set objRecT = Server.CreateObject ("ADODB.Recordset")
    
    If Len(cEmail) > 6 Then
	SQLText = "Select TotalCost From " & Session("ShowTransactions") & " Where (ClientEmail = '" & cEmail & "') And (OrderDate > #" & tDate & "#) And (Quote <> 'Yes' Or Quote Is Null) And (Items Like '%S___/%' Or Items Like '%W___/%') And (Items Not Like '%S212/%')"
	
	objRecT.Open SQLText, transcationPath, 3, 3, 0
	If Not (objRecT.EOF Or objRecT.BOF) Then
	    objRecT.MoveFirst
	    Do 
		clientCost = clientCost + objRecT("TotalCost")
		objRecT.MoveNext
		If objRecT.EOF Then Exit Do
	    Loop
	End If
	objRecT.Close
    End If
    
    If clientCost = 0 And Len(cPostal) > 0 Then
	'SQLText = "Select TotalCost From Transaction Where (ClientPostalCode = '" & cPostal & "') And (OrderDate > #" & tDate & "#) And (Quote <> 'Yes' Or Quote Is Null) And (Items Like '%S___/%') And (Items Not Like '%S212/%')"
	SQLText = "Select TotalCost From " & Session("ShowTransactions") & " Where (ClientPostalCode = '" & cPostal & "') And (ClientCompany = '" & comName & "') And (OrderDate > #" & tDate & "#) And (Quote <> 'Yes' Or Quote Is Null) And (Items Like '%S___/%' Or Items Like '%W___/%') And (Items Not Like '%S212/%')"
	
        objRecT.Open SQLText, transcationPath, 3, 3, 0
        If Not (objRecT.EOF Or objRecT.BOF) Then
            objRecT.MoveFirst
            Do While Not objRecT.EOF
                 officeCost = officeCost + objRecT("TotalCost")
                 objRecT.MoveNext
            Loop
        End If
        objRecT.Close
    End If
    
    If clientCost = 0 And officeCost = 0 And Len(comName) > 0 Then
	SQLText = "Select TotalCost From " & Session("ShowTransactions") & " Where (ClientCompany = '" & comName & "') And (OrderDate > #" & tDate & "#) And (Quote <> 'Yes' Or Quote Is Null) And (Items Like '%S___/%' Or Items Like '%W___/%') And (Items Not Like '%S212/%')"
	
        objRecT.Open SQLText, transcationPath, 3, 3, 0
        If Not (objRecT.EOF Or objRecT.BOF) Then
            objRecT.MoveFirst
            Do While Not objRecT.EOF
                 companyCost = companyCost + objRecT("TotalCost")
                 objRecT.MoveNext
            Loop
        End If
        objRecT.Close
    End If
    
    Set objRecT = Nothing    
  
    Dim Costs(2)
    Costs(0) = clientCost
    Costs(1) = officeCost
    Costs(2) = companyCost
    
    ClientCheck = Costs

End Function
%>