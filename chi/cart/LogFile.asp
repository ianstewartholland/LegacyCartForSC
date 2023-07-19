
<%

' Activation service logging helper.

' Write a log message to the logfile.
Public Function WritePivotalLog(strMessage)
	Dim logfile: logfile = "C:\WWWUsers\LocalUser\chi\computationalhydraulics_com\databases\pivotal.log"
	Dim f, fs
	
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	Set f = fs.OpenTextFile(logfile, 8, true)
	
	f.WriteLine(Date & " " & Time & ": " & strMessage)
	f.Close()
	Set f = Nothing
	Set fs = Nothing
	
End Function

'
' other functions that may be shared by the cart
'
' Remove any non-numeric characters from the
' string "x" and return.
Public Function MakeNumeric(x)
				MakeNumeric = ""
				i = 1
				do while i <= Len(x)
								if isNumeric(Mid(x, i, 1)) then
												MakeNumeric = MakeNumeric & Mid(x, i, 1)
								end if
								i = i + 1					
				loop
End Function



' Hide the credit card number in the familiar
' xxxx-xxxx-xxxx-1234 format.
' Also implemented in checkout.asp. Fix re: DRY.
' FIXMEAMEX
Public Function HideCreditCardNumber(ccNumber)
				Dim cleanedNumber
				
				' First remove any non-numeric characters
				cleanedNumber = MakeNumeric(ccNumber)
				
				If Len(cleanedNumber) < 12 Then
								cleanedNumber = ""
								Exit Function
				End If
				
				HideCreditCardNumber = "xxxx-xxxx-xxxx-"
				' Take the last 3 (Amex) or 4 (Visa, MC, Discover) characters of the credit card.
				HideCreditCardNumber = HideCreditCardNumber & Right(cleanedNumber, Len(cleanedNumber) - 12)

End Function

'
' display dropdown list for state/province
' this is also used by address.asp - same logic but changed appearance
' 
Sub DisplayProvinceDropdown(country, province)
    ' define array list to store the states or province
    Dim stateList()

    Select Case country
        Case "USA"
            ReDim stateList(60)
	    stateList(0) = "Please select..."
            stateList(1) = "Alabama - AL"
	    stateList(2) = "Alaska - AK"
	    stateList(3) = "Arizona - AZ"
	    stateList(4) = "Arkansas - AR"
	    stateList(5) = "California - CA"
	    stateList(6) = "Colorado - CO"
	    stateList(7) = "Connecticut - CT"
	    stateList(8) = "Delaware - DE"
	    stateList(9) = "District of Columbia - DC"
	    stateList(10) = "Florida - FL"
	    stateList(11) = "Georgia - GA"
	    stateList(12) = "Hawaii - HI"
	    stateList(13) = "Idaho - ID"
	    stateList(14) = "Illinois - IL"
	    stateList(15) = "Indiana - IN"
	    stateList(16) = "Iowa - IA"
	    stateList(17) = "Kansas - KS"
	    stateList(18) = "Kentucky - KY"
	    stateList(19) = "Louisiana - LA"
	    stateList(20) = "Maine - ME"
	    stateList(21) = "Maryland - MD"
	    stateList(22) = "Massachusetts - MA"
	    stateList(23) = "Michigan - MI"
	    stateList(24) = "Minnesota - MN"
	    stateList(25) = "Mississippi - MS"
	    stateList(26) = "Missouri - MO"
	    stateList(27) = "Montana - MT"
	    stateList(28) = "Nebraska - NE"
	    stateList(29) = "Nevada - NV"
	    stateList(30) = "New Hampshire - NH"
	    stateList(31) = "New Jersey - NJ"
	    stateList(32) = "New Mexico - NM"
	    stateList(33) = "New York - NY"
	    stateList(34) = "North Carolina - NC"
	    stateList(35) = "North Dakota - ND"
	    stateList(36) = "Ohio - OH"
	    stateList(37) = "Oklahoma - OK"
	    stateList(38) = "Oregon - OR"
	    stateList(39) = "Pennsylvania - PA"
	    stateList(40) = "Rhode Island - RI"
	    stateList(41) = "South Carolina - SC"
	    stateList(42) = "South Dakota - SD"
	    stateList(43) = "Tennessee - TN"
	    stateList(44) = "Texas - TX"
	    stateList(45) = "Utah - UT"
	    stateList(46) = "Vermont - VT"
	    stateList(47) = "Virginia - VA"
	    stateList(48) = "Washington - WA"
	    stateList(49) = "West Virginia - WV"
	    stateList(50) = "Wisconsin - WI"
	    stateList(51) = "Wyoming - WY"
	    stateList(52) = "American Samoa - AS"
	    stateList(53) = "Guam - GU"
	    stateList(54) = "Federated States of Micronesia - FM"
	    stateList(55) = "Marshall Islands - MH"
	    stateList(56) = "Northern Mariana Islands - MP"
	    stateList(57) = "Palau - PW"
	    stateList(58) = "Puerto Rico - PR"
	    stateList(59) = "Virgin Islands - WY"
            
            Response.Write "<div class='row'><div class='col-sm-3'>State *</div><div class='col-sm-9'><select size='1' name='Province' id='Province' >"
            
            For t = 0 to 59
                If Right(stateList(t),2) = province or Left(stateList(t), Len(province))  = province Then
                    Response.Write "<option selected>" & stateList(t) & "</option>"
                Else
                    Response.Write "<option>" & stateList(t) & "</option>"
                End If
	    Next
            Response.Write "</select></div></div>"

        Case "Canada"
            ReDim stateList(14)
	    stateList(0) = "Please select..."
	    stateList(1) = "Alberta - AB"
	    stateList(2) = "British Columbia - BC"
	    stateList(3) = "Manitoba - MB"
	    stateList(4) = "New Brunswick - NB"
	    stateList(5) = "Newfoundland - NL"
	    stateList(6) = "Northwest Territories - NT"
	    stateList(7) = "Nova Scotia - NS"
	    stateList(8) = "Nunavut - NU"
	    stateList(9) = "Ontario - ON"
	    stateList(10) = "Prince Edward Island - PE"
	    stateList(11) = "Quebec - QC"
	    stateList(12) = "Saskatchewan - SK"
	    stateList(13) = "Yukon Territory - YT"
            
            Response.Write "<div class='row'><div class='col-sm-3'>Province *</div><div class='col-sm-9'><select size='1' name='Province' id='Province'>"
            For t = 0 to 13
                If Right(stateList(t),2) = province or Left(stateList(t), Len(province))  = province Then
                    Response.Write "<option selected value = '" & Right(stateList(t), 2) & "'>" & stateList(t) & "</option>"
                Else
                    Response.Write "<option value = '" & Right(stateList(t), 2) & "'>" & stateList(t) & "</option>"
                End If
	    Next
            
            Response.Write "</select></div></div>"
        Case Else
            Response.Write "<div class='row'><div class='col-sm-3'>Province/State (eg. Region, District) *</div><div class='col-sm-9'>"
            Response.Write "<input type='text' name='Province' id='Province' size='20' value='" &  province & "' maxlength='15'>"
            Response.Write "</div></div>"
    End Select

End Sub

'
' email extension filter
'
Function EmailCheckFailed(email)
    EmailCheckFailed = False
    ' define an array of email extension to look for
    Dim blockList(4)
    blockList(1) = "@gmail."
    blockList(2) = "@hotmail."
    blockList(3) = "@yahoo."
    blockList(4) = "@ymail."
    
    For t = 1 to UBound(blockList)
        If InStr(email, blockList(t)) > 0 Then
            EmailCheckFailed = True
            Exit For
        End If
    Next

End Function

'
' province look up function - render user's input of province and make a best guess
'
Function ValidateProvince(country, province)
    Dim stateList()

    ValidProvince = False
    orgProvince = province
    province = UCase(province)
    ' according to country, check province
    Select Case UCase(country)
	Case "CANADA"
            ReDim stateList(13)
	    stateList(1) = "Alberta - AB"
	    stateList(2) = "British Columbia - BC"
	    stateList(3) = "Manitoba - MB"
	    stateList(4) = "New Brunswick - NB"
	    stateList(5) = "Newfoundland - NL"
	    stateList(6) = "Northwest Territories - NT"
	    stateList(7) = "Nova Scotia - NS"
	    stateList(8) = "Nunavut - NU"
	    stateList(9) = "Ontario - ON"
	    stateList(10) = "Prince Edward Island - PE"
	    stateList(11) = "Quebec - QC"
	    stateList(12) = "Saskatchewan - SK"
	    stateList(13) = "Yukon Territory - YT"
	Case "USA"
            ReDim stateList(59)
            stateList(1) = "Alabama - AL"
	    stateList(2) = "Alaska - AK"
	    stateList(3) = "Arizona - AZ"
	    stateList(4) = "Arkansas - AR"
	    stateList(5) = "California - CA"
	    stateList(6) = "Colorado - CO"
	    stateList(7) = "Connecticut - CT"
	    stateList(8) = "Delaware - DE"
	    stateList(9) = "District of Columbia - DC"
	    stateList(10) = "Florida - FL"
	    stateList(11) = "Georgia - GA"
	    stateList(12) = "Hawaii - HI"
	    stateList(13) = "Idaho - ID"
	    stateList(14) = "Illinois - IL"
	    stateList(15) = "Indiana - IN"
	    stateList(16) = "Iowa - IA"
	    stateList(17) = "Kansas - KS"
	    stateList(18) = "Kentucky - KY"
	    stateList(19) = "Louisiana - LA"
	    stateList(20) = "Maine - ME"
	    stateList(21) = "Maryland - MD"
	    stateList(22) = "Massachusetts - MA"
	    stateList(23) = "Michigan - MI"
	    stateList(24) = "Minnesota - MN"
	    stateList(25) = "Mississippi - MS"
	    stateList(26) = "Missouri - MO"
	    stateList(27) = "Montana - MT"
	    stateList(28) = "Nebraska - NE"
	    stateList(29) = "Nevada - NV"
	    stateList(30) = "New Hampshire - NH"
	    stateList(31) = "New Jersey - NJ"
	    stateList(32) = "New Mexico - NM"
	    stateList(33) = "New York - NY"
	    stateList(34) = "North Carolina - NC"
	    stateList(35) = "North Dakota - ND"
	    stateList(36) = "Ohio - OH"
	    stateList(37) = "Oklahoma - OK"
	    stateList(38) = "Oregon - OR"
	    stateList(39) = "Pennsylvania - PA"
	    stateList(40) = "Rhode Island - RI"
	    stateList(41) = "South Carolina - SC"
	    stateList(42) = "South Dakota - SD"
	    stateList(43) = "Tennessee - TN"
	    stateList(44) = "Texas - TX"
	    stateList(45) = "Utah - UT"
	    stateList(46) = "Vermont - VT"
	    stateList(47) = "Virginia - VA"
	    stateList(48) = "Washington - WA"
	    stateList(49) = "West Virginia - WV"
	    stateList(50) = "Wisconsin - WI"
	    stateList(51) = "Wyoming - WY"
	    stateList(52) = "American Samoa - AS"
	    stateList(53) = "Guam - GU"
	    stateList(54) = "Federated States of Micronesia - FM"
	    stateList(55) = "Marshall Islands - MH"
	    stateList(56) = "Northern Mariana Islands - MP"
	    stateList(57) = "Palau - PW"
	    stateList(58) = "Puerto Rico - PR"
	    stateList(59) = "Virgin Islands - WY"
    End Select
  
    If Len(province) < 2 Then
	' invalid
    ElseIf Len(province) = 2 Then
	For t = 1 to UBound(stateList)
	    If Right(stateList(t), 2) = province Then 
		ValidProvince = True
		Exit For
	    Else
		' incorrect province abbr
		ValidProvince = False
	    End If
	Next
    ElseIf Len(province) < 5 Then
	For t = 1 to UBound(stateList)
	    If Left(UCase(stateList(t)), Len(province)) = Left(province, Len(province)) Then
		ValidProvince = True
		province =  Right(stateList(t), 2)
		Exit For
	    Else
		ValidProvince = False
	    End If
	Next
    Else	' not in abbr nor short - search the first five characters
	For t = 1 to UBound(stateList)
	    If Left(UCase(stateList(t)), 5) = Left(province, 5) Then
		If Left(stateList(t), 5) <> "North" And Left(stateList(t), 5) <> "South" Then
		    ValidProvince = True
		    province = Right(stateList(t), 2)
		    Exit For
		ElseIf Len(province) > 6 Then
		    If Left(UCase(stateList(t)), 7) = Left(province, 7) Then
			ValidProvince = True
			province =  Right(stateList(t), 2)
			Exit For
		    Else
			ValidProvince = False
		    End If
		End If
	    Else
		ValidProvince = False
	    End If
	Next
    End If

    ' restore original value
    If Not ValidProvince Then province = orgProvince 
    
    Dim outs(2)
    outs(1) = province
    outs(2) = ValidProvince
    
    ValidateProvince = outs
    
End Function

'
' basic email address validation for @ and . after @ with at least one character
'
Function ValidateEmail(email)
    ValidateEmail = False
    If InStr(email, "@") > 1 And InStr(Mid(email, InStr(email, "@") + 2), ".") > 0 Then ValidateEmail = True

End Function

%>

 