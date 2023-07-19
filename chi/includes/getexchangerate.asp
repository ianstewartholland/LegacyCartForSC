
<!--#include file ="dbpaths.inc"-->
<%
'
' get the exchange rate for the orders
' for now, assume the last exchange rate of an order is the latest exchange rate
'
Function GetExchangeRate(CurrencyAbrev)

    If CurrencyAbrev = "US" Then
	'Set objRecRate = Server.CreateObject ("ADODB.Recordset")	
	'objRecRate.Open "Transaction", transcationPath, 3, 3, 2
	'objRecRate.MoveLast
	'Do until objRecRate.BOF    
	'    If Left(objRecRate("TransactionNumber"), 1) <> "Q" And objRecRate("CurrencyExchange") > 0 Then
	'	GetExchangeRate = objRecRate("CurrencyExchange")
	'	Exit Do
	'    End If
	'    objRecRate.MovePrevious
	'Loop
	'objRecRate.Close
	'Set objRecRate = Nothing
	
	' for now, use 1
	GetExchangeRate = 1
    Else    
	GetExchangeRate = 0
    End If

End Function
%>