<!--#include virtual ="/chi/includes/dbpaths.inc"-->
<%
'
' we do not ever load the catalog in sessions - for workshop and conference, read from the planner database; for else read from catalog database
' for the costs, we assume the exchange rate is 1 for US and canadian dollors
'
Sub AddItem(ItemStr, ItemNum)
    Dim objRecSR, conn
    ItemStr = UCase(ItemStr)
    Response.Write itemstr
    ' check whether this item already exists in the session if yes add quantity if no add item
    For i = 1 To Session("NumCartItems")
	If Session("CartItem(" & i & ")") = ItemStr Then
	    Session("CartItemQuantity(" & i & ")") = Session("CartItemQuantity(" & i & ")") + ItemNum

	    ' double check if R999X and reset it's value
	    If ItemStr = "R184" Or ItemStr = "R242" Then

		For j = 1 To Session("NumCartItems")
		    If Session("CartItem(" & j & ")") = "R999X" Then
			Session("CartItemCost(" & j & ")") = Session("CartItemCost(" & j & ")") - Session("CartItemCost(" & i & ")") * 0.2
			Exit For
		    End If
		Next
	    End If

	    Exit Sub
	End If
   Next

    t = Session("NumCartItems") + 1
    Session("NumCartItems") = t
    Session("AskWorkshopInterests") = False

    Select Case Left(ItemStr, 1)
	Case "W", "P"
	    ' check whether this is an advanced workshop
	    isAdvanced = False
        IsInCanada = False
        cleanItemStr = ItemStr
	    'If Len(ItemStr) > 4 And Left(ItemStr, 1) = "W" Then
        If Len(ItemStr) > 4 Then
            If Not IsNumeric(Right(ItemStr, 1)) Then
		        isAdvanced = True
                cleanItemStr = Left(ItemStr, Len(ItemStr) - 1)
	        End If
	    End If

	    ' set connection object here
	    set conn=Server.CreateObject("ADODB.Connection")
	    conn.Provider="Microsoft.Jet.OLEDB.4.0"
	    conn.Open workshopPath

	    Set objRecWk = Server.CreateObject ("ADODB.Recordset")
	    'objRecWk.Open "Workshops", workshopPath, 3, 3, 2
	    'objRecWk.Filter = "WorkshopID = '" & Left(ItemStr, 4) & "'"
	    'objRecWk.Open "Select * from [Workshops] Where [WorkshopID] = '" & Left(ItemStr, 4) & "'", conn
        objRecWk.Open "Select * from [Workshops] Where [WorkshopID] = '" & cleanItemStr & "'", conn
	    objRecWk.MoveFirst

    If Not IsNull(objRecWk("AskInterests")) And objRecWk("AskInterests")  Then Session("AskWorkshopInterests") = True

	    If objRecWk("Country") = "Web Training" Then
		Session("CartItem(" & t & ")") = ItemStr
		Session("CartItemDescription(" & t & ")") = objRecWk("WorkshopName") ' "PCSWMM & SWMM5 self paced " & WebTrainingItem '& " " & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration")) 
        If objRecWk("IsSpecialty") Then Session("CartItemDescription(" & t & ")") = Replace(Session("CartItemDescription(" & t & ")"), "workshop", "specialty workshop")   

If Not Session("CartItem(" & t & ")") = "W1000" Then
    If Right(ItemStr, 1) = "B" Then 
        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & " (" & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate"))) & ")"
    ElseIf Right(ItemStr, 1) = "A" Or objRecWk("Duration") = 1 Then 
        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & " (" & MonthName(Month(objRecWk("StartDate"))) & " " & Day(objRecWk("StartDate")) & ", " & Year(objRecWk("StartDate")) & ")"
    Else
        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & " (" & MonthName(Month(objRecWk("StartDate"))) & " " & Day(objRecWk("StartDate")) & "-" &  Day(DateAdd("d",objRecWk("Duration")-1,objRecWk("StartDate"))) & ", " & Year(objRecWk("StartDate")) & ")"
    End If 
Else
    Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") 
End If
        If Right(ItemStr, 1) = "A"  Or Right(ItemStr, 1) = "B" Then
		    Session("CartItemCost(" & t & ")") = objRecWk("AdvancedCost")
        Else
		    Session("CartItemCost(" & t & ")") = objRecWk("Cost")
        End If
		Session("CartItemQuantity(" & t & ")") = ItemNum
		Session("WebWorkshop") = True
	    Else
		Session("CartItem(" & t & ")") = ItemStr
        If LCase(objRecWk("Country")) = "canada" Then IsInCanada = True

		If Left(ItemStr, 1) = "P" Then
		    Session("CartItemDescription(" & t & ")") = FormatStartDate(objRecWk("StartDate"), 5) & ", " & objRecWk("City")
		    Session("CartItemCost(" & t & ")") = objRecWk("Cost")
		Else
		    If Not isAdvanced Then
                If IsInCanada Then
			        If objRecWk("Province") <> "" Then
			            Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 " & objRecWk("City") & ", " & objRecWk("Province") & " workshop & time limited software " & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration"))
			        Else
			            Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 " & objRecWk("City") & ", " & objRecWk("Country") & " workshop & time limited software " & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration"))
			        End If
                Else
			        If objRecWk("Province") <> "" Then
			            Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with training - " & objRecWk("City") & ", " & objRecWk("Province") & " " & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration"))
			        Else
			            Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with training - " & objRecWk("City") & ", " & objRecWk("Country") & " " & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration"))
			        End If
                End If
			Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") 

			Session("CartItemCost(" & t & ")") = objRecWk("Cost")
		    Else
			' assume we only allow they advance on the second (last) day
    
                If IsInCanada Then
			        If objRecWk("Province") <> "" Then
			            ' Duration = 3 for SA and the last day is the advanced day
			            If objRecWk("Duration") = 3 Then
				            If Right(ItemStr, 1) = "C" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 advanced 1 day "
				            If Right(ItemStr, 1) = "C" Then Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & objRecWk("City") & ", " & objRecWk("Province") & " workshop & time limited software " & MonthName(Month(DateAdd("d",2,objRecWk("StartDate")))) & " " & Day(DateAdd("d",2,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",2,objRecWk("StartDate")))
				            If Right(ItemStr, 1) = "B" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 " & objRecWk("City") & ", " & objRecWk("Province") & " workshop & time limited software " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
				            If Right(ItemStr, 1) = "A" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 " & objRecWk("City") & ", " & objRecWk("Province") & " workshop & time limited software " & MonthName(Month(objRecWk("StartDate"))) & " " & Day(objRecWk("StartDate")) & ", " & Year(objRecWk("StartDate"))
			            Else
                                        If LCase(Right(ItemStr, 3)) = "lid" Then
                                            Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 LIDs 1 day " & objRecWk("City") & ", " & objRecWk("Province") & " workshop & time limited software " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
                                            Session("CartItem(" & t & ")") = Left(ItemStr, 5)
                                        Else
                                            Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 advanced and 2D 1 day " & objRecWk("City") & ", " & objRecWk("Province") & " workshop & time limited software " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
                                        End If
			            End If
			        Else
			            If objRecWk("Duration") = 3 Then
				        If Right(ItemStr, 1) = "C" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 advanced 1 day "
				        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & objRecWk("City") & ", " & objRecWk("Country") & " workshop & time limited software " & MonthName(Month(DateAdd("d",2,objRecWk("StartDate")))) & " " & Day(DateAdd("d",2,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",2,objRecWk("StartDate")))
				        If Right(ItemStr, 1) = "B" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5" & objRecWk("City") & ", " & objRecWk("Country") & " workshop & time limited software " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
				        If Right(ItemStr, 1) = "A" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5" & objRecWk("City") & ", " & objRecWk("Country") & " workshop & time limited software " & MonthName(Month(objRecWk("StartDate"))) & " " & Day(objRecWk("StartDate")) & ", " & Year(objRecWk("StartDate"))
			            Else
                                        Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 advanced 1 day " & objRecWk("City") & ", " & objRecWk("Country") & " workshop & time limited software " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
			            End If
			        End If
                Else    ' outside canada
                    If objRecWk("Province") <> "" Then
			            ' Duration = 3 for SA and the last day is the advanced day
			            If objRecWk("Duration") = 3 Then
				            If Right(ItemStr, 1) = "C" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with advanced 1 day training - "
				            If Right(ItemStr, 1) = "C" Then Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & objRecWk("City") & ", " & objRecWk("Province") & " " & MonthName(Month(DateAdd("d",2,objRecWk("StartDate")))) & " " & Day(DateAdd("d",2,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",2,objRecWk("StartDate")))
				            If Right(ItemStr, 1) = "B" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with training - " & objRecWk("City") & ", " & objRecWk("Province") & " " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
				            If Right(ItemStr, 1) = "A" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with training - " & objRecWk("City") & ", " & objRecWk("Province") & " " & MonthName(Month(objRecWk("StartDate"))) & " " & Day(objRecWk("StartDate")) & ", " & Year(objRecWk("StartDate"))
			            Else
                                        If LCase(Right(ItemStr, 3)) = "lid" Then
                                            Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with LIDs 1 day training - " & objRecWk("City") & ", " & objRecWk("Province") & " " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
                                            Session("CartItem(" & t & ")") = Left(ItemStr, 5)
                                        Else
                                            Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with advanced and 2D 1 day training - " & objRecWk("City") & ", " & objRecWk("Province") & " " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
                                        End If
			            End If
			        Else
			            If objRecWk("Duration") = 3 Then
				        If Right(ItemStr, 1) = "C" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with advanced 1 day training - "
				        Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")") & objRecWk("City") & ", " & objRecWk("Country") & " " & MonthName(Month(DateAdd("d",2,objRecWk("StartDate")))) & " " & Day(DateAdd("d",2,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",2,objRecWk("StartDate")))
				        If Right(ItemStr, 1) = "B" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with training - " & objRecWk("City") & ", " & objRecWk("Country") & " " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
				        If Right(ItemStr, 1) = "A" Then Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with training - " & objRecWk("City") & ", " & objRecWk("Country") & " " & MonthName(Month(objRecWk("StartDate"))) & " " & Day(objRecWk("StartDate")) & ", " & Year(objRecWk("StartDate"))
			            Else
                                        Session("CartItemDescription(" & t & ")") = "PCSWMM & SWMM5 software limited subscription with advanced 1 day training - " & objRecWk("City") & ", " & objRecWk("Country") & " " & MonthName(Month(DateAdd("d",1,objRecWk("StartDate")))) & " " & Day(DateAdd("d",1,objRecWk("StartDate"))) & ", " & Year(DateAdd("d",1,objRecWk("StartDate")))
			            End If
			        End If
            End If
			Session("CartItemDescription(" & t & ")") = Session("CartItemDescription(" & t & ")")

			Session("CartItemCost(" & t & ")") = objRecWk("AdvancedCost")
		    End If
    
            ' for specialty workshop
            If objRecWk("IsSpecialty") Then
                Session("CartItemDescription(" & t & ")") = Replace(Session("CartItemDescription(" & t & ")"),"PCSWMM & SWMM5", objRecWk("WorkshopName") & " - Specialty Course, ")
            End If
		End If

		Session("CartItemQuantity(" & t & ")") = ItemNum
	    End If

	    objRecWk.Close
	    Set objRecWk = Nothing
	    conn.Close
	    Set conn = Nothing	
    Case "S", "R", "T", "A", "U"
	    strSQL = "Select * from [Catalog] Where [ItemNo] = '" & ItemStr & "'"
	    ' set connection object here
	    set conn=Server.CreateObject("ADODB.Connection")
	    conn.Provider="Microsoft.Jet.OLEDB.4.0"
	    conn.Open catalogPath
	    set objRecSR = Server.CreateObject ("ADODB.Recordset")
	    objRecSR.Open strSQL, conn

	    If Not objRecSR.EOF Then
		objRecSR.MoveFirst

		Session("CartItem(" & t & ")") = ItemStr
		Session("CartItemDescription(" & t & ")") = objRecSR("Description")
		Session("CartItemCost(" & t & ")") = objRecSR("CADCost")
		Session("CartItemQuantity(" & t & ")") = ItemNum
	    End If
	    objRecSR.Close
	    Set objRecSR = Nothing
	    conn.Close
	    Set conn = Nothing

	Case "C"
	    Set objRecWk = Server.CreateObject ("ADODB.Recordset")
	    'objRecWk.Open "Conferences", conferencePath, 3, 3, 2
	    'objRecWk.Filter = "ConferenceID = 'C0" & Right(ItemStr, 2) & "'"

	    ' set connection object here
	    set conn=Server.CreateObject("ADODB.Connection")
	    conn.Provider="Microsoft.Jet.OLEDB.4.0"
	    conn.Open conferencePath

	    objRecWk.Open "Select * from [Conferences] Where ConferenceID = 'C0" & Right(ItemStr, 2) & "'", conn

	    objRecWk.MoveFirst

	    Session("CartItem(" & t & ")") = ItemStr

	    Select Case Left(ItemStr, 2)
		Case "C0"
		    Session("CartItemDescription(" & t & ")") = "Conference registration (" & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration")) & ")"
		    Session("CartItemCost(" & t & ")") = objRecWk("ConferenceRegistrationFee")
		Case "C2"
		    Session("CartItemDescription(" & t & ")") = "Presenter registration (" & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration")) & ")"
		    Session("CartItemCost(" & t & ")") = objRecWk("PresenterRegistrationFee")
		Case "C3"
		    Session("CartItemDescription(" & t & ")") = "Exhibitor registration (" & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration")) & ")"
		    Session("CartItemCost(" & t & ")") = objRecWk("ExhibitorRegistrationFee")
		Case "C4"
		    Session("CartItemDescription(" & t & ")") = "Add'l exhibitor registration (" & FormatStartDate(objRecWk("StartDate"), objRecWk("Duration")) & ")"
		    Session("CartItemCost(" & t & ")") = objRecWk("AddlExhibitorRegistrationFee")
	    End Select

	    Session("CartItemQuantity(" & t & ")") = ItemNum

	    objRecWk.Close
	    Set objRecWk = Nothing
	    conn.Close
	    Set conn = Nothing
    End Select

End Sub

'
' this function is also used by addfromcatalog.asp
'
Function FormatStartDate(StartDate, duration)
    ' check whether this add day to the next month
    newDate = DateAdd("d", duration - 1, StartDate)

    If duration = 2 Then
	If Month(StartDate) = Month(newDate) Then
	    FormatStartDate = MonthName(Month(StartDate)) & " " & Day(StartDate) & " & " & Day(newDate) & ", " & Year(StartDate)
	Else
	    FormatStartDate = MonthName(Month(StartDate)) & " " & Day(StartDate) & " & " & MonthName(Month(newDate)) & " " & Day(newDate) & ", " & Year(StartDate)
	End If
    ElseIf duration > 2 And duration < 5 Then
	If Month(StartDate) = Month(newDate) Then
	    FormatStartDate = MonthName(Month(StartDate)) & " " & Day(StartDate) & " - " & Day(newDate) & ", " & Year(StartDate)
	Else
	    FormatStartDate = MonthName(Month(StartDate)) & " " & Day(StartDate) & " - " & MonthName(Month(newDate)) & " " & Day(newDate) & ", " & Year(StartDate)
	End If
    ElseIf duration = 5 Then
	FormatStartDate = MonthName(Month(StartDate)) & " " & Day(StartDate) & ", " & Year(StartDate)
    Else
	FormatStartDate = MonthName(Month(StartDate)) & " " & Day(StartDate) & ", " & Year(StartDate) & " (" & duration
	If duration = 1 Then
	    FormatStartDate = FormatStartDate & " day)"
	Else
	    FormatStartDate = FormatStartDate & " days)"
	End If
    End If

End Function

%>