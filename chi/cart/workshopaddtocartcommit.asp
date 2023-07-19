<!--#include virtual ="/chi/includes/AddItem.asp"-->
<%
	'check to process form
	If Session("WorkshopRegistrationInProgress") <> "ON" Then
		'Session("Location") = Request.Form("Location")
		'Session("Software") = Request.Form("Software")
		'Session("EducationalDiscount") = Request.Form("EducationalDiscount")
		Location = Request.Form("Location")
		Software = Request.Form("Software")
		EducationalDiscount = Request.Form("EducationalDiscount")
		Session("WorkshopRegistrationInProgress") = "ON"
		IndiaDiscount = Request.Form("IndiaDiscount")
		R184 = Request.Form("R184")
		R242 = Request.Form("R242")
	End If
response.write Location & "<br>"
	Session("WorkshopRegistrationInProgress") = "OFF"

	'remove any existing workshop, software or workshop book items from cart
	Do
		For t = 1 to Session("NumCartItems")
			Select Case Left(Session("CartItem(" & t & ")"),1)
				Case "W", "S", "D"
					Exit For
			End Select
			Select Case Session("CartItem(" & t & ")")
				Case "R242", "R184", "R999X", ""
					Exit For
			End Select
		Next
		If t > Session("NumCartItems") Then Exit Do
		'item found
		Session("NumCartItems") = Session("NumCartItems") - 1
		For n = t to Session("NumCartItems")
			Session("CartItem(" & n & ")") = Session("CartItem(" & n + 1 & ")")
			Session("CartItemDescription(" & n & ")") = Session("CartItemDescription(" & n + 1 & ")")
			Session("CartItemCost(" & n & ")") = Session("CartItemCost(" & n + 1 & ")")
			Session("CartItemQuantity(" & n & ")") = Session("CartItemQuantity(" & n + 1 & ")")
		Next
		Session("CartItem(" & n & ")") = ""
		Session("CartItemDescription(" & n & ")") = ""
		Session("CartItemCost(" & n & ")") = 0
		Session("CartItemQuantity(" & n & ")") = 0
	Loop

	' check to see whether this is an indian workshop
	Dim indianWorkshops(5)
	indianWorkshops(1) = "W658"
	indianWorkshops(2) = "W659"
	indianWorkshops(3) = "W660"
	indianWorkshops(4) = "W661"
	indianWorkshops(5) = "W583"
	isIndia = False



	' check if SA workshop
	' keep in mind the check code here is cheap. you should read from database see if they got postponed or changed number instead of below
	Dim saWorkshops(3)
	saWorkshops(1) = "W1192"
	saWorkshops(2) = "W1135"
	saWorkshops(3) = "W1136"

	isSa = False
	startdate = ""

	For t = 1 to 5
	    If InStr(Location, indianWorkshops(t)) > 0 Then
		isIndia = True
		Exit For
	    End If
	Next

	For t = 1 to UBound(saWorkshops)
	    If InStr(Location, saWorkshops(t)) > 0 Then
		isSa = True
		startdate = Mid(Location, 6)
		Exit For
	    End If
	Next

	If InStr(Location, "W719") > 0 Then
		isUK = True
		startdate = Mid(Location, 5)
		If Len(Location) > 4 Then
	    	Location = Left(Location, 4)
	    End If
	End If

	' filter out the sa location
	If isSa And Len(Location) > 4 Then
	    Location = Left(Location, 5)
	End If

	SaDay = 0
	discountPrice = 0

	If isIndia Then
	    ' assume there is no advanced workshop in india
	    AddItem Location, 1
	    If IndiaDiscount = "ON" Then
		t = Session("NumCartItems") + 1
		Session("NumCartItems") = t
		Session("CartItem(" & t & ")") = Location & "X"
		Session("CartItemDescription(" & t & ")") = "Student/Government agency discount"
		Session("CartItemCost(" & t & ")") = - 85
		' special discount for india
		Session("CartItemQuantity(" & t & ")") = 1
	    End If

	    If Software = "ON" then
		AddItem "S202", 1

		t = Session("NumCartItems") + 1
		Session("NumCartItems") = t
		Session("CartItem(" & t & ")") = "S202X"
		Session("CartItemDescription(" & t & ")") = "PCSWMM 2011 Professional workshop discount"
		If Not isAdvanced Then
		    Session("CartItemCost(" & t & ")") = - 200
		Else
		    Session("CartItemCost(" & t & ")") = - 100
		End If
		Session("CartItemQuantity(" & t & ")") = 1

	    End If
	ElseIf isSa Then
	    ' for sa workshops, check 4 checkbox
	    'If Request.Form("LocationAll") = "Yes" Or (Request.Form("LocationA") = "Yes" And Request.Form("LocationB") = "Yes" And Request.Form("LocationC") = "Yes") Or (Request.Form("LocationAll") <> "Yes" And Request.Form("LocationA") <> "Yes" And Request.Form("LocationB") <> "Yes" And Request.Form("LocationC") <> "Yes") Then
    If Request.Form("LocationAll") = "Yes" Or (Request.Form("LocationAB") = "Yes" And Request.Form("LocationC") = "Yes") Or (Request.Form("LocationAll") <> "Yes" And Request.Form("LocationAB") <> "Yes" And Request.Form("LocationC") <> "Yes") Then
		    AddItem Location, 1
		    pricing = Session("CartItemCost(" & t & ")")
		    discountPrice = pricing
		    SaDay = 3
            ' 2 day SA workshop
	    Else 
            If Request.Form("LocationAB") = "Yes" Then 'And Request.Form("LocationB") = "Yes" And (Location = "W771" Or Location = "W769") Then    
		        AddItem Location & "A", 1
		        Session("CartItemCost(" & t & ")") =  265
    
                ' item desc
                If InStr(Session("CartItemDescription(" & t & ")"), "PCSWMM & SWMM5 ") = 0 Then
                    Session("CartItemDescription(" & t & ")") = Replace(Session("CartItemDescription(" & t & ")"), "PCSWMM & SWMM5", "PCSWMM & SWMM5 ")
                End If

		        AddItem Location & "B", 1
		        Session("CartItemCost(" & t & ")") =  265
    
                ' item desc
                If InStr(Session("CartItemDescription(" & t & ")"), "PCSWMM & SWMM5 ") = 0 Then
                    Session("CartItemDescription(" & t & ")") = Replace(Session("CartItemDescription(" & t & ")"), "PCSWMM & SWMM5", "PCSWMM & SWMM5 ")
                End If

		        pricing = 530
		        discountPrice = 530
		        SaDay = 2
                

            Else
		If Request.Form("LocationC") = "Yes" Then
		    AddItem Location & "C", 1

		    pricing = Session("CartItemCost(" & t & ")")
		    discountPrice = pricing
		    SaDay = SaDay + 1
		End If
		'If Request.Form("LocationB") = "Yes" Then
		 '   AddItem Location & "B", 1
		  '  If SaDay = 1 Then Session("CartItemCost(" & t & ")") = 210
		   ' pricing = pricing +  Session("CartItemCost(" & t & ")")
		    'discountPrice = pricing
		   ' SaDay = SaDay + 1
		'End If
'		If Request.Form("LocationC") = "Yes" Then
'		    AddItem Location & "C", 1
'		    If SaDay = 1 Then Session("CartItemCost(" & t & ")") = 210
'		    pricing = pricing + Session("CartItemCost(" & t & ")")
'		    discountPrice = pricing
'		    SaDay = SaDay + 1
'		End If
	    End If
        End If


	    If EducationalDiscount = "ON" Then
		t = Session("NumCartItems") + 1
		Session("NumCartItems") = t
		Session("CartItem(" & t & ")") = Location & "G"
		Session("CartItemDescription(" & t & ")") = "Educational grant"
		Session("CartItemCost(" & t & ")") = - pricing * 0.5
		discountPrice = discountPrice - pricing * 0.5

		Session("CartItemQuantity(" & t & ")") = 1
	    End If
	    ' early bird 15% discount till 30 days before
	    ' If DateDiff("d", Date, CDate(#1/30/11#)) >= 0 Then EarlyBirdDiscountFactor = 0.1
	    ' remember to check date before parse it
        If (InStr(location, "1192") > 0 And DateDiff("d",Date,"1-Sep-2022") > 0) Then
    		t = Session("NumCartItems") + 1
		    Session("NumCartItems") = t
		    Session("CartItem(" & t & ")") = Location & "X"
		    Session("CartItemDescription(" & t & ")") = "Early bird discount"
		    Session("CartItemCost(" & t & ")") = - discountPrice * 0.2

		    Session("CartItemQuantity(" & t & ")") = 1
	    ElseIf (IsDate(startdate) And DateDiff("d", Date, CDate(startdate)) >= 30 And Location <> "W1018") Then
		    t = Session("NumCartItems") + 1
		    Session("NumCartItems") = t
		    Session("CartItem(" & t & ")") = Location & "X"
		    Session("CartItemDescription(" & t & ")") = "Early bird discount" 
		    Session("CartItemCost(" & t & ")") = - discountPrice * 0.15

		    Session("CartItemQuantity(" & t & ")") = 1
	    End If
	    ' add software based on days
	    If Software = "ON" then
		AddItem "S202", 1

		t = Session("NumCartItems") + 1
		Session("NumCartItems") = t
		Session("CartItem(" & t & ")") = "S202X"
		Session("CartItemDescription(" & t & ")") = "PCSWMM 2011 Professional South Africa workshop discount"

		' discount is added based on days
		If SaDay = 3 Then
		    Session("CartItemCost(" & t & ")") = - 300
		ElseIf SaDay = 2 Then
		    Session("CartItemCost(" & t & ")") = - 200
		ElseIf SaDay = 1 Then
		    Session("CartItemCost(" & t & ")") = - 100
		End If

		Session("CartItemQuantity(" & t & ")") = 1
	    End If
	ElseIf isUK Then
	    ' Uk london workshop in 2013
	    If Request.Form("LocationAll") = "Yes" Or (Request.Form("LocationA") = "Yes" And Request.Form("LocationB") = "Yes") Or (Request.Form("LocationAll") <> "Yes" And Request.Form("LocationA") <> "Yes" And Request.Form("LocationB") <> "Yes") Then
			pricing = - 930
	    Else
			If Request.Form("LocationA") = "Yes" Then
				Location = Location & "A"
			End If

			If Request.Form("LocationB") = "Yes" Then
				Location = Location & "B"
			End If

			pricing = - 615
		End If

		AddItem Location, 1

		' cheap hack for day 1 incorrect date
		If Request.Form("LocationA") = "Yes" And Request.Form("LocationB") <> "Yes" And Request.Form("LocationAll") <> "Yes" Then
			t = Session("NumCartItems")
			Session("CartItemDescription(" & t & ")") = "Advanced 1 day London, United Kingdom workshop & time limited software November 18, 2013"
		End If

	    If EducationalDiscount = "ON" Then
				t = Session("NumCartItems") + 1
				Session("NumCartItems") = t
				Session("CartItem(" & t & ")") = Location & "X"
				Session("CartItemDescription(" & t & ")") = "London United Kingdom workshop student discount"

				Session("CartItemCost(" & t & ")") = pricing

				Session("CartItemQuantity(" & t & ")") = 1
	    End If
	Else
response.write "here" & Location
        ' get startdate and location
        startdate = ""
        If InStr(Location, "-") > 1 Then
            startdate = Mid(Location, InStr(Location, "-") + 1)
            Location = Left(Location, InStr(Location, "-") - 1)
        End If

	    'add workshop registration to cart
	    Session("WebWorkshop") = False

    ' check location has multiple days
    If Len(Location) = 7 Then
        WorkshopItem = Left(Location, 6)
        AddItem WorkshopItem, 1
        WorkshopItem = Left(Location, 5) & Right(Location, 1)
        AddItem WorkshopItem, 1
	    pricing = Session("CartItemCost(" & t & ")") + Session("CartItemCost(" & t-1 & ")")
    ElseIf Location = "W1126" Then
        AddItem Location, 1
        t = Session("NumCartItems")
        Session("CartItemCost(" & t & ")") = 595*3
	    pricing = Session("CartItemCost(" & t & ")")
    Else
        AddItem Location, 1
	    pricing = Session("CartItemCost(" & t & ")")
    End If

        If Location = "W1156" Then
            t = Session("NumCartItems")
			Session("CartItemDescription(" & t & ")") = "PCSWMM + SWMM5 pre-conference workshop (February 28 - March 1, 2022)"
        End If 

	    ' check whether this is an advanced workshop for the software discount
	    isAdvanced = False
	    If Len(Location) > 4 And Left(Location, 1) = "W" Then isAdvanced = True


	    'check to add manuals to self-paced workshop
	    'If Session("WebWorkshop") Then
	    '	AddItem "R233", 1
	    '	AddItem "R184", 1
	    'End If

	    'add educational discount to cart

	    'If Session("EducationalDiscount") = "ON" Then
	    If EducationalDiscount = "ON" Then
		    If Session("WebWorkshop") Then
		        t = Session("NumCartItems") + 1
		        Session("NumCartItems") = t
		        Session("CartItem(" & t & ")") = Location & "G"
		        Session("CartItemDescription(" & t & ")") = "Web workshop educational grant"
		        Session("CartItemCost(" & t & ")") = - Round(pricing * 0.3 + 1)
		        Session("CartItemQuantity(" & t & ")") = 1
		    Else
		        t = Session("NumCartItems") + 1
		        Session("NumCartItems") = t
		        ' location may contains the LIDs
		        If InStr(Location, "LID") > 0 Then Location = Replace(Location, "LID", "")

            If Ucase(Right(Location,1)) = "A" Or Ucase(Right(Location,1)) = "B" Or Ucase(Right(Location,1)) = "C" Then
                WorkshopItem = Left(Location, Len(Location) - 1)
                If Ucase(Right(WorkshopItem,1)) = "A" Or Ucase(Right(WorkshopItem,1)) = "B" Or Ucase(Right(WorkshopItem,1)) = "C" Then
                    WorkshopItem = Left(WorkshopItem, Len(WorkshopItem) - 1)
                End If
                Session("CartItem(" & t & ")") = WorkshopItem & "G"
			Else
			    Session("CartItem(" & t & ")") = Location & "G"
			End If

		        'Session("CartItem(" & t & ")") = Left(Location, 4) & "G"
		        Session("CartItemDescription(" & t & ")") = "Educational grant"
		        Session("CartItemCost(" & t & ")") = - Round(pricing * 0.3 + 1)
		        Session("CartItemQuantity(" & t & ")") = 1
		    End If
            
            pricing = pricing - Round(pricing * 0.3 + 1)
	    End If

response.write "here" & startdate
    ' early bird
    If Len(startdate) > 0 And IsDate(startdate) And DateDiff("d", Date, CDate(startdate)) >= 26 And Not Session("WebWorkshop") And Location <> "W1083" Then
        t = Session("NumCartItems") + 1
		Session("NumCartItems") = t
		' location may contains the LIDs
        ' sep 14, 2017 - cannot count on item length to trim as workshop number exceeds 1000
		'Session("CartItem(" & t & ")") = Left(Location, 4) & "X"

            If Ucase(Right(Location,1)) = "A" Or Ucase(Right(Location,1)) = "B" Or Ucase(Right(Location,1)) = "C" Then
                WorkshopItem = Left(Location, Len(Location) - 1)
                If Ucase(Right(WorkshopItem,1)) = "A" Or Ucase(Right(WorkshopItem,1)) = "B" Or Ucase(Right(WorkshopItem,1)) = "C" Then
                    WorkshopItem = Left(WorkshopItem, Len(WorkshopItem) - 1)
                End If
                WorkshopItem = WorkshopItem & "X"
                Session("CartItem(" & t & ")") = WorkshopItem
			Else
			    Session("CartItem(" & t & ")") = Location & "X"
			End If

    If Session("WebWorkshop") Then
		Session("CartItemDescription(" & t & ")") = "Web workshop early registration discount"
    Else
		Session("CartItemDescription(" & t & ")") = "Early bird discount"
    End If
		Session("CartItemCost(" & t & ")") = - Round(pricing * 0.1/5)*5
		Session("CartItemQuantity(" & t & ")") = 1
    End If
	    'add software to cart
	    'If Session("Software") = "ON" then
	    'If Software = "ON" then
		'    AddItem "S202", 1
		'    If Session("WebWorkshop") Then
		'	t = Session("NumCartItems") + 1
		'	Session("NumCartItems") = t
		'	Session("CartItem(" & t & ")") = "S202X"
		'	Session("CartItemDescription(" & t & ")") = "PCSWMM 2011 Professional web workshop discount"
		'	Session("CartItemCost(" & t & ")") = - 100
		'	Session("CartItemQuantity(" & t & ")") = 1
		'    Else
		'	t = Session("NumCartItems") + 1
		'	Session("NumCartItems") = t
		'	Session("CartItem(" & t & ")") = "S202X"
		'	Session("CartItemDescription(" & t & ")") = "PCSWMM 2011 Professional workshop discount"
		'	If Not isAdvanced Then
		'	    Session("CartItemCost(" & t & ")") = - 200
		'	Else
		'	    Session("CartItemCost(" & t & ")") = - 100
		'	End If
		'	Session("CartItemQuantity(" & t & ")") = 1
		'    End If
	    'End If

	    ' add books for an advanced workshop - should be outside of this branch but be safe for the indians
	    'If (R184 = "ON" Or R242 = "ON") And Len(Location) > 4 Then
		'price = 0
        '
		'If R184 = "ON" Then
		'    AddItem "R184", 1
		'    price = price + Session("CartItemCost(" & t & ")")
		'End If
        '
		'If R242 = "ON" Then
		'    AddItem "R242", 1
		'    price = price + Session("CartItemCost(" & t & ")")
		'End If
        '
		'' add discount for the two books - 20%
		't = Session("NumCartItems") + 1
		'Session("NumCartItems") = t
		'Session("CartItem(" & t & ")") = "R999X"
		'Session("CartItemDescription(" & t & ")") = "Advanced workshop materials discount"
		'Session("CartItemCost(" & t & ")") = - 0.2 * price
		'Session("CartItemQuantity(" & t & ")") = 1

	    'End If

	End If ' is india

        ' Latornell 2019
        If Location = "W1083" Then
            t = Session("NumCartItems") 
            Session("CartItemDescription(" & t & ")") = "Pre-Latornell Workshop - Planning and executing stormwater retrofit studies using EPA SWMM5&PCSWMM, November 18, 2019 (1 day)"
        End If

'    If Location = "W1018" Then
'        t = Session("NumCartItems")
'		 Session("CartItemDescription(" & t & ")") = "PCSWMM Professional 2D single user subscription (12 months) with training workshop - Hanoi, Vietnam August 9 & 10, 2018"
'    End If 


	Response.Redirect "cart.asp"

%>




