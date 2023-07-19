<%
'
' this is to display for the backend only
'
Sub DisplayProvinceDropdownBackEnd(country, province)
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
            
            Response.Write "<tr><td width='30%' bgcolor='#FEFFD7'><font face='Arial' size='2'>State</font></td><td width='70%' bgcolor='#FEFFD7'><select size='1' name='Province' id='Province' >"
            
            For t = 0 to 59
                If Right(stateList(t),2) = province or Left(stateList(t), Len(province))  = province Then
                    Response.Write "<option selected>" & stateList(t) & "</option>"
                Else
                    Response.Write "<option>" & stateList(t) & "</option>"
                End If
	    Next
            Response.Write "</select></td></tr>"

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
            
            Response.Write "<tr><td width='30%' bgcolor='#FEFFD7'><font face='Arial' size='2'>Province</font></td><td width='70%' bgcolor='#FEFFD7'><select size='1' name='Province' id='Province'>"
            For t = 0 to 13
                If Right(stateList(t),2) = province or Left(stateList(t), Len(province))  = province Then
                    Response.Write "<option selected value = '" & Right(stateList(t), 2) & "'>" & stateList(t) & "</option>"
                Else
                    Response.Write "<option value = '" & Right(stateList(t), 2) & "'>" & stateList(t) & "</option>"
                End If
	    Next
            
            Response.Write "</select></td></tr>"
        Case Else
            Response.Write "<tr><td width='30%' bgcolor='#FEFFD7'><font face='Arial' size='2'>Province/State</font></td><td width='70%' bgcolor='#FEFFD7'>"
            Response.Write "<input type='text' name='Province' id='Province' size='20' value='" &  province & "' maxlength='15'>"
            Response.Write "</td></tr>"
    End Select

End Sub
%>
