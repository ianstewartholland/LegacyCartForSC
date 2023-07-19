﻿<!--#include virtual ="/chi/includes/AddItem.asp"-->
<%

	Dim EarlyBirdDiscount 
        
        'check to process form
	If Session("ConfSemRegistrationInProgress") <> "ON" Then
            Workshop = Request.Form("Workshop")
            Software = Request.Form("Software")
            Software2 = Request.Form("Software2")
            'Session("Manuals") = Request.Form("Manuals")
            Conference = Request.Form("Conference")
            Conference1 = Request.Form("Conference1")
            Conference2 = Request.Form("Conference2")
            'ConferenceType = Request.Form("ConferenceType")
            Presenter = Request.Form("Presenter")
            AdditionalExhibitors = Request.Form("AdditionalExhibitors")
            EducationalDiscount = Request.Form("EducationalDiscount")
            's add check box for the conference dinner
            ConferenceDinner = Request.Form("ConferenceDinner")
    
    
            IsCADPricing = Request.Form("IsPriceCAD")
            FieldTrip = Request.Form("FieldTrip")

    PPT = Request.Form("PPT")
    ResearchProject = Request.Form("ResearchProject")
            
            WorkshopDay1 = Request.Form("WorkshopDay1")
            WorkshopDay2 = Request.Form("WorkshopDay2")
            'get number of workshop days client is registering for
            'check if at least one day or conference has been selected
            If Workshop <> "ON" And WorkshopDay1 <> "ON" And WorkshopDay2 <> "ON" And Conference <> "ON" And Conference1 <> "ON" And Conference2 <> "ON" Then
                    Response.Redirect("http://www.icwmm.org")
            End If
            Session("ConfSemRegistrationInProgress") = "ON"
	End If
	
	Session("ConfSemRegistrationInProgress") = "OFF"

	'remove any existing conference/workshop items from cart
	Do
            For t = 1 to Session("NumCartItems")
                    If Left(Session("CartItem(" & t & ")"),1) = "W" Or Left(Session("CartItem(" & t & ")"),1) = "C" Or Left(Session("CartItem(" & t & ")"),1) = "D" Or Left(Session("CartItem(" & t & ")"),4) = "S220" Or Session("CartItem(" & t & ")") = "" Then Exit For
            Next
            If t > Session("NumCartItems") Then Exit Do
            'software item found
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
'
' notes on editting this page
' edit date for the early bird discount
' edit workshop/software/conference item number and time item 
' if there is an extra item, eg C520 the conference dinner, that's another story - add the code yourself
'
        EarlyBirdDiscountFactor = 0
        'If Date < DateSerial(2012,1,22) Then
        '    EarlyBirdDiscountFactor = 0.1
        'Else
        '    EarlyBirdDiscountFactor = 0
        'End If
	' date format "month/day/year", discount applies on this date - remember to change this date on the checkout.asp as well otherwise next year won't be good
	If DateDiff("d", Date, CDate(#12/31/22#)) >= 0 Then EarlyBirdDiscountFactor = 0.1
	WorkshopItem = "W1196"
    WorkshopDate = "2/27/23"
	SoftwareItem = "S220"
	SoftwareItem2 = "S222"
	ConferenceItem = "32"	'eg. "C021"
	TimeItem = " (March 1-2, 2023)"	' will append to presenter/exhibitor registration description
    strDay1 = "(March 1, 2023)"
    strDay2 = "(March 2, 2023)"
    ' still need to edit date for dinner item below don't forget
    conPrice = 345
    con1DayPrice = 0
    If DateDiff("d", CDate(#2/1/23#), Date) >= 0 Then
        conPrice = 395
    End If
    HasWorkshop = False
	
	'add workshop registration to cart
	If Workshop = "ON" then
            AddItem WorkshopItem, 1
	    HasWorkshop = True
	    ' change description
        If WorkshopItem = "W1196" Then
            Session("CartItemDescription(" & t & ")") = "PCSWMM + EPA SWMM5 modeling workshop (February 27-28, 2023)"
        End If
	    'Session("CartItemDescription(" & t & ")") = "Advanced 1 day Toronto, ON workshop & time limited software February 21, 2012 (includes workbook)"
	        WorkshopCost = Session("CartItemCost(" & t & ")")
            If EducationalDiscount = "ON" Then
                t = Session("NumCartItems") + 1
                Session("NumCartItems") = t
                Session("CartItem(" & t & ")") = WorkshopItem & "G"
                Session("CartItemDescription(" & t & ")") = "Virtual live training educational grant" 'educational grant
                Session("CartItemCost(" & t & ")") = - Round((WorkshopCost * 0.3)/5)*5
                Session("CartItemQuantity(" & t & ")") = 1
                WorkshopCost = WorkshopCost - Round((WorkshopCost * 0.3)/5)*5
            End If
            'WorkshopCost = Session("CartItemCost(" & t & ")") * Session("CartItemQuantity(" & t & ")")'response.write "here" &  (Left(WorkshopItem, 5) = "W1141")

    If DateDiff("d", Date, CDate(WorkshopDate)) >= 26 Then
        response.Write "here"
    End If 


            If Not (Left(WorkshopItem, 5) = "W1196") Then
                If DateDiff("d", Date, CDate(WorkshopDate)) >= 26 Then
                    t = Session("NumCartItems") + 1
                    Session("NumCartItems") = t
                    Session("CartItem(" & t & ")") = WorkshopItem & "X"
                    Session("CartItemDescription(" & t & ")") = "Workshop early registration discount" 
                    Session("CartItemCost(" & t & ")") = - Round((WorkshopCost * 0.1)/ 5)*5
                    Session("CartItemQuantity(" & t & ")") = 1
                End If
            End If
	End If

	If WorkshopDay1 = "ON" And Workshop <> "ON" then
            AddItem WorkshopItem & "A", 1
	    HasWorkshop = True
	    ' change description
        If WorkshopItem = "W1196" Then
            Session("CartItemDescription(" & t & ")") = "PCSWMM + EPA SWMM5 modeling workshop (February 27, 2023)"
            Session("CartItemCost(" & t & ")") = 595
        End If
	    'Session("CartItemDescription(" & t & ")") = "Advanced 1 day Toronto, ON workshop & time limited software February 21, 2012 (includes workbook)"
	        WorkshopCost = Session("CartItemCost(" & t & ")")
            If EducationalDiscount = "ON" Then
                t = Session("NumCartItems") + 1
                Session("NumCartItems") = t
                Session("CartItem(" & t & ")") = WorkshopItem & "G"
                Session("CartItemDescription(" & t & ")") = "Virtual live training educational grant" 'educational grant
                Session("CartItemCost(" & t & ")") = - Round(WorkshopCost * 0.3 / 5)*5
                Session("CartItemQuantity(" & t & ")") = 1
                WorkshopCost = WorkshopCost - Round(WorkshopCost * 0.3 + 1)
            End If
            'WorkshopCost = Session("CartItemCost(" & t & ")") * Session("CartItemQuantity(" & t & ")")
    
            If Not (Left(WorkshopItem, 5) = "W1196") Then
                If DateDiff("d", Date, CDate(WorkshopDate)) >= 26 Then
                    t = Session("NumCartItems") + 1
                    Session("NumCartItems") = t
                    Session("CartItem(" & t & ")") = WorkshopItem & "X"
                    Session("CartItemDescription(" & t & ")") = "Workshop early registration discount"
                    'Session("CartItemCost(" & t & ")") = - Round(WorkshopCost * 0.1 / 5)*5
                    Session("CartItemCost(" & t & ")") = -100
                    Session("CartItemQuantity(" & t & ")") = 1
                End If
            End If
            'If Software = "ON" Then
            '    AddItem SoftwareItem, 1
            '    t = Session("NumCartItems") + 1
            '    Session("NumCartItems") = t
            '    Session("CartItem(" & t & ")") = SoftwareItem & "X"
            '    Session("CartItemDescription(" & t & ")") = "PCSWMM 2012 Professional workshop discount"
            '    Session("CartItemCost(" & t & ")") = - 200
            '    Session("CartItemQuantity(" & t & ")") = 1
            'End If
	End If

	If WorkshopDay2 = "ON" And Workshop <> "ON" then
            AddItem WorkshopItem & "B", 1
	    HasWorkshop = True
	    ' change description
        If WorkshopItem = "W1196" Then
            Session("CartItemDescription(" & t & ")") = "PCSWMM + EPA SWMM5 modeling workshop (February 28, 2023)"
            Session("CartItemCost(" & t & ")") = 595
        End If
	    'Session("CartItemDescription(" & t & ")") = "Advanced 1 day Toronto, ON workshop & time limited software February 21, 2012 (includes workbook)"
	        WorkshopCost = Session("CartItemCost(" & t & ")")
            If EducationalDiscount = "ON" Then
                t = Session("NumCartItems") + 1
                Session("NumCartItems") = t
                Session("CartItem(" & t & ")") = WorkshopItem & "G"
                Session("CartItemDescription(" & t & ")") = "Virtual live training educational grant" 'educational grant
                Session("CartItemCost(" & t & ")") = - Round(WorkshopCost * 0.3 / 5)*5
                Session("CartItemQuantity(" & t & ")") = 1
                WorkshopCost = WorkshopCost - Round(WorkshopCost * 0.3 + 1)
            End If
            'WorkshopCost = Session("CartItemCost(" & t & ")") * Session("CartItemQuantity(" & t & ")")
    
            If Not (Left(WorkshopItem, 5) = "W1196") Then
                If DateDiff("d", Date, CDate(WorkshopDate)) >= 26 Then
                    t = Session("NumCartItems") + 1
                    Session("NumCartItems") = t
                    Session("CartItem(" & t & ")") = WorkshopItem & "X"
                    Session("CartItemDescription(" & t & ")") = "Workshop early registration discount"
                    'Session("CartItemCost(" & t & ")") = - Round(WorkshopCost * 0.1 / 5)*5
                    Session("CartItemCost(" & t & ")") = -100
                    Session("CartItemQuantity(" & t & ")") = 1
                End If
            End If
            'If Software = "ON" Then
            '    AddItem SoftwareItem, 1
            '    t = Session("NumCartItems") + 1
            '    Session("NumCartItems") = t
            '    Session("CartItem(" & t & ")") = SoftwareItem & "X"
            '    Session("CartItemDescription(" & t & ")") = "PCSWMM 2012 Professional workshop discount"
            '    Session("CartItemCost(" & t & ")") = - 200
            '    Session("CartItemQuantity(" & t & ")") = 1
            'End If
	End If
        If Not HasWorkshop Then Session("QuoteQuery") = "Conferences"

            If Software = "ON" Then
                AddItem SoftwareItem, 1
                't = Session("NumCartItems") + 1
                'Session("NumCartItems") = t
                'Session("CartItem(" & t & ")") = SoftwareItem & "X"
                'Session("CartItemDescription(" & t & ")") = "PCSWMM 2012 Professional workshop discount"
                'Session("CartItemCost(" & t & ")") = - 200
                'Session("CartItemQuantity(" & t & ")") = 1
            End If
            If Software2 = "ON" Then
                AddItem SoftwareItem2, 1
                't = Session("NumCartItems") + 1
                'Session("NumCartItems") = t
                'Session("CartItem(" & t & ")") = SoftwareItem & "X"
                'Session("CartItemDescription(" & t & ")") = "PCSWMM 2012 Professional workshop discount"
                'Session("CartItemCost(" & t & ")") = - 200
                'Session("CartItemQuantity(" & t & ")") = 1
            End If
	'add conference registration to cart
	earlyTotal = 0
    registerConference = False
    price = conPrice
    If Conference = "ON" Or (Conference1 = "ON" And Conference2 = "ON") then

        ' two days
        registerConference = True
        t = Session("NumCartItems") + 1
		Session("NumCartItems") = t
		Session("CartItem(" & t & ")") = "C0" & ConferenceItem 
		Session("CartItemDescription(" & t & ")") = "Conference registration " & TimeItem
		Session("CartItemCost(" & t & ")") = conPrice
		Session("CartItemQuantity(" & t & ")") = 1
        earlyTotal = earlyTotal + Session("CartItemCost(" & t & ")")
    ElseIf Conference1 = "ON" Then
        ' day 1
        registerConference = True
        price = con1DayPrice 'Round(((conPrice + 5)*0.65)/5)*5
        ' if late reg
       ' If conPrice > 500 Then price = 395
        t = Session("NumCartItems") + 1
		Session("NumCartItems") = t
		Session("CartItem(" & t & ")") = "C0" & ConferenceItem & "A"
		Session("CartItemDescription(" & t & ")") = "Conference registration " & strDay1 
		Session("CartItemCost(" & t & ")") = price
		Session("CartItemQuantity(" & t & ")") = 1
        earlyTotal = earlyTotal + Session("CartItemCost(" & t & ")")
    ElseIf Conference2 = "ON" Then
        ' day 2
        registerConference = True
        price = con1DayPrice 'Round((conPrice*0.65)/5)*5
        ' if late reg
       ' If conPrice > 500 Then price = 395
        t = Session("NumCartItems") + 1
		Session("NumCartItems") = t
		Session("CartItem(" & t & ")") = "C0" & ConferenceItem & "B"
		Session("CartItemDescription(" & t & ")") = "Conference registration " & strDay2 
		Session("CartItemCost(" & t & ")") = price
		Session("CartItemQuantity(" & t & ")") = 1
        earlyTotal = earlyTotal + Session("CartItemCost(" & t & ")")
    Else
        ' didn't register conference
    End If


' to test
'Response.Write price
    If registerConference Then
        If EarlyBirdDiscountFactor > 0 Then

            ' calculate early bird amount, nov 2018
            earlyAmount = 0
            For t = 1 to Session("NumCartItems")
                If Left(Session("CartItem(" & t & ")"), 1) = "C" Then 
                    earlyAmount = earlyAmount + Session("CartItemCost(" & t & ")")
                End If
            Next

            t = Session("NumCartItems") + 1 
            Session("NumCartItems") = t
            Session("CartItem(" & t & ")") = "C0" & ConferenceItem & "X"
            Session("CartItemDescription(" & t & ")") = "Conference early registration discount"
    ' custom early bird rate
            'q = 30
            'If price > 400 Then q = 50
            'Session("CartItemCost(" & t & ")") = - Round((earlyTotal - q) * EarlyBirdDiscountFactor) 
            'Session("CartItemCost(" & t & ")") = - q 
            'Session("CartItemCost(" & t & ")") = - Round(earlyAmount * EarlyBirdDiscountFactor / 5)*5
            Session("CartItemCost(" & t & ")") = -50
            Session("CartItemQuantity(" & t & ")") = 1

            price = price - 50
        End If

        If EducationalDiscount = "ON" Then
			t = Session("NumCartItems") + 1
			Session("NumCartItems") = t
			Session("CartItem(" & t & ")") = "C0" & ConferenceItem & "G"
			Session("CartItemDescription(" & t & ")") = "Conference educational grant" 'educational grant
            Session("CartItemCost(" & t & ")") = - Round(price * 0.3 / 5)*5
			Session("CartItemQuantity(" & t & ")") = 1
            earlyTotal = earlyTotal + Session("CartItemCost(" & t & ")")
        End If

        If Presenter = "ON" Then 's both attendee and presenter
            's add presenter with no cost
            t = Session("NumCartItems") + 1
            Session("NumCartItems") = t
            Session("CartItem(" & t & ")") = "C2" & ConferenceItem
            Session("CartItemDescription(" & t & ")") = "Presenter registration" & TimeItem
            Session("CartItemCost(" & t & ")") = 0.00
            Session("CartItemQuantity(" & t & ")") = 1
        End If

        's add conference dinner
        If ConferenceDinner = "ON" Then
		    t = Session("NumCartItems") + 1
            Session("NumCartItems") = t
            Session("CartItem(" & t & ")") = "C5" & ConferenceItem
            Session("CartItemDescription(" & t & ")") = "Conference dinner – Wednesday February 24, 2021" 
            Session("CartItemCost(" & t & ")") = 0.00
            Session("CartItemQuantity(" & t & ")") = 1
	    End If

'       If FieldTrip = "ON" Then
'		    t = Session("NumCartItems") + 1
'           Session("NumCartItems") = t
'           Session("CartItem(" & t & ")") = "C6" & ConferenceItem
'           Session("CartItemDescription(" & t & ")") = "Field trip – Friday March 3, 2017" 
'           Session("CartItemCost(" & t & ")") = 85.00
'           Session("CartItemQuantity(" & t & ")") = 1
'       End If
'   
'           ' add tshirt item
'           If Request.Form("TShirt") = "ON" And Not Len(Request.Form("Size")) > 3 Then 
'			    t = Session("NumCartItems") + 1
'			    Session("NumCartItems") = t
'			    Session("CartItem(" & t & ")") = "C7" & ConferenceItem 
'			    Session("CartItemDescription(" & t & ")") = "Conference T-Shirt (" & Request.Form'("size") & ")" ' add t-shirt size
'               Session("CartItemCost(" & t & ")") = 0
'			    Session("CartItemQuantity(" & t & ")") = 1
'            End If
    End If ' registerConference

    ' update to US pricing
'    If IsCADPricing  = "OFF" Then
'        For t = 1 to Session("NumCartItems")
'            ' enable the condition below nov 2018
'            If Instr(Session("CartItem(" & t & ")"), WorkshopItem) < 1 Then 
'                Session("CartItemCost(" & t & ")") = Round((Session("CartItemCost(" & t & ")") / 1.25)/5)*5
'            End If
'        Next
'        Session("ChangePricing") = True
'    End If

    ' take care of PPT and RP
    permissionStr = ""
    If PPT = "ON" Then 
        permissionStr = "UsePPT-YES"
    Else
        permissionStr = "UsePPT-NO"
    End If
    If ResearchProject = "ON" Then 
        permissionStr = permissionStr & "|RP-YES"
    Else
        permissionStr = permissionStr & "|RP-NO"
    End If 
    Session("ConferencePermission") = permissionStr

	Response.Redirect "cart.asp"

%>





 