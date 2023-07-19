<%
' not the best place for this 
' function to format license datetime with eg. expire in 3 weeks - logic borrowed from crm (to keep things consist)
Function LicenseDateStr(str)
    licDate = CDate(str)

    days = DateDiff("d", licDate, Date)
  
    If days < 0 Then
	days = days * -1
	showAgo = False
    Else
	showAgo = True
    End If
    
    ' assemble the string
    If days < 22 Then
	If days < 2 Then
	    str = GetDate(licDate) & " (" & days & " day)" 
	Else
	    str = GetDate(licDate) & " (" & days & " days)" 
	End If
    ElseIf days > 21 And days < 61 Then
    
	days = Round(days / 7)
	
	If days < 2 Then
	    str = GetDate(licDate) & " (" & days & " week)" 
	Else
	    str = GetDate(licDate) & " (" & days & " weeks)" 
	End If
    ElseIf days > 60 And days < 30 * 24 Then
     
	days = Round(days / 30.4167)
	
	If days < 2 Then
	    str = GetDate(licDate) & " (" & days & " month)" 
	Else
	    str = GetDate(licDate) & " (" & days & " months)" 
	End If
    Else

	days = Round(days / 365.25)

	If days < 2 Then
	    str = GetDate(licDate) & " (" & days & " year)" 
	Else
	    str = GetDate(licDate) & " (" & days & " years)" 
	End If
    End If
    
    If showAgo Then
	str = Replace(str, ")", " ago)")
    Else
	str = Replace(str, ")", " from now)")
    End If

    LicenseDateStr = str
End Function

' get date format like Nov 11, 2011
Function GetDate(myDate)
    GetDate = MonthName(Month(myDate), true) & " " & Day(myDate) & ", " & Year(myDate) 
End Function

'
'function to handle text field with special tags eg. http:// into html
'copyright - Rob
'
Function InsertTags(Field)
    ' remove extra line breaks from the end of string
    Do While Len(Field) >= 2
	Field = Trim(Field)
	If Right(Field, 2) = vbCrLf Then
	    Field = Left(Field, Len(Field) - 2)
	Else
	    Exit Do
	End If
    Loop

    NewField = " " & Trim(Field) & " "
    
    'replace hard returns (and insert spaces)
    NewField = Replace(NewField, Chr(13) & Chr(10), " <br> ")

    'tag ^ h3
    i = InStr(NewField , "^")
    Do While i <> 0
    	j = i
    	k = FindEndIgnorePeriod(i + 1, NewField)
        Part1 = Left(NewField, j - 1)
        strURL = Mid(NewField, j + 1, k - j)
        Part2 = Mid(NewField,k)
        strURLTags = "<h3>" & strURL & "</h3>"
        NewField = Part1 & strURLTags & Part2
        i = InStr(i + Len(strURLTags), NewField, "^")
    Loop

    'tag hyperlinks
    i = InStr(NewField , "http://")
    Do While i <> 0
    	j = i
    	k = FindEndIgnorePeriod(i + 1, NewField)
        Part1 = Left(NewField, j - 1)
        strURL = Mid(NewField, j, k - j)
        Part2 = Mid(NewField,k)
        strURLTags = "<a href='" & strURL & "' Target='_blank' >" & strURL & "</a>"
        NewField = Part1 & strURLTags & Part2
        i = InStr(i + Len(strURLTags), NewField, "http://")
    Loop

    'tag order numbers and license numbers
If Not isLocked Then    
    i = InStr(NewField , "#")
    

    Do While i <> 0
    	j = i
    	k = FindEnd(i + 1, NewField)
        If k > i + 1 Then
            Part1 = Left(NewField, j - 1)
            strLink = Mid(NewField, j+1, k - j - 1)
            Part2 = Mid(NewField,k)
            If IsNumeric(strLink) Then
                If Len(strLink) = 3 Then
                    strLink = "<a href='/chi/licenses/enterpriselicense.asp?EnterpriseLicense=" & strLink & "'>#" & strLink & "</a>"
                ElseIf Len(strLink) <= 5 Then
                    strLink = "<a href='/chi/orders/order.asp?transaction=" & strLink & "'>#" & strLink & "</a>"
                ElseIf Len(strLink) = 6 Then
		    strLink = "<a href='/chi/licenses/license.asp?license=" & strLink & "'>#" & strLink & "</a>"
                End If
	    ElseIf Len(strLink) = 6 And Left(strLink, 1) = "Q" Then
		strLink = "<a href='/chi/orders/order.asp?transaction=" & strLink & "'>#" & strLink & "</a>"
            End If
            NewField = Part1 & strLink & Part2
            i = i + Len(strLink)
        End If
        i = InStr(i + 1, NewField, "#")
    Loop
End If

    'tag M:\ links
    i = InStr(1, NewField , " m:\", 1)
    Do While i <> 0
    	j = i
        'look for space after extension (doesn't work if more than one ".")
    	k = Instr(i, NewField, ".")
    	k = Instr(k + 1, NewField, " ")
        Part1 = Left(NewField, j - 1)
        strURL = Mid(NewField, j, k - j)
        Part2 = Mid(NewField,k)
        strURLTags = "<a href='" & strURL & "' Target='_blank' >" & strURL & "</a>"
        NewField = Part1 & strURLTags & Part2
        i = InStr(i + Len(strURLTags), NewField, "m:\", 1)
    Loop
    
    'tag emails
    i = InStr(NewField , "@")
    Do While i <> 0
    	j = InstrRev(NewField, " ", i) + 1
    	k = Instr(i, NewField, " ")
        If i > j + 1 and k > i + 1 Then
            Part1 = Left(NewField, j - 1)
            strEmail = Mid(NewField, j, k - j)
            Part2 = Mid(NewField,k)
            strEmailTags = "<a href='mailto:" & strEmail & "'>" & strEmail & "</a>"
            NewField = Part1 & strEmailTags & Part2
            i = i + Len(strEmailTags)
        End If
        i = InStr(i + 1, NewField, "@")
    Loop

    'tag !! - make red
    i = InStr(NewField , "!!")
    Do While i <> 0
    	j = i
    	k = Instr(i + 2, NewField, "!!")
        if k = 0 then k = len(NewField)
        If k > i + 2 Then
            Part1 = Left(NewField, j - 1)
            strImportant = Mid(NewField, j+2, k - j - 2)
            Part2 = Mid(NewField,k+2)
            strImportantTags = "<font color='red'>" & strImportant & "</font>"
            NewField = Part1 & strImportantTags & Part2
            i = i + Len(strImportantTags)
        End If
        i = InStr(i + 1, NewField, "!!")
    Loop

    'tag ! - make red
    i = InStr(NewField , "!")
    Do While i <> 0
    	j = i
    	k = FindEnd(i + 1, NewField)
        If k > i + 1 Then
            Part1 = Left(NewField, j - 1)
            strImportant = Mid(NewField, j+1, k - j)
            Part2 = Mid(NewField,k)
            strImportantTags = "<font color='red'>" & strImportant & "</font>"
            NewField = Part1 & strImportantTags & Part2
            i = i + Len(strImportantTags)
        End If
        i = InStr(i + 1, NewField, "!")
    Loop
    
    'tag "important" - make red
    'i = InStr(1, NewField , "important", 1)
    'Do While i <> 0
    '	j = i
    '	k = FindEnd(i + 1, NewField)
     '   If k > i + 1 Then
      '      Part1 = Left(NewField, j - 1)
       '     strImportant = Mid(NewField, j, k - j)
        '    Part2 = Mid(NewField,k)
         '   strImportantTags = "<font color='red'>" & strImportant & "</font>"
          '  NewField = Part1 & strImportantTags & Part2
           ' i = i + Len(strImportantTags)
        'End If
        'i = InStr(i + 1, NewField, "important", 1)
'    Loop
    
    'tag "follow up" - make red
'    i = InStr(1, NewField , "follow up", 1)
 '   Do While i <> 0
  '  	j = i
'    	k = FindEnd(i + 8, NewField)
'        If k > i + 1 Then
 '           Part1 = Left(NewField, j - 1)
  '          strImportant = Mid(NewField, j, k - j)
   '         Part2 = Mid(NewField, k)
    '        strImportantTags = "<font color='red'>" & strImportant & "</font>"
     '       NewField = Part1 & strImportantTags & Part2
      '      i = i + Len(strImportantTags) 
       ' End If
        'i = InStr(i + 1, NewField, "follow up", 1)
    'Loop
    
    'tag "followed up" - make red
'    i = InStr(1, NewField , "followed up", 1)
 '   Do While i <> 0
  '  	j = i
   ' 	k = FindEnd(i + 10, NewField)
    '    If k > i + 1 Then
     '       Part1 = Left(NewField, j - 1)
      '      strImportant = Mid(NewField, j, k - j)
       '     Part2 = Mid(NewField, k)
        '    strImportantTags = "<font color='red'>" & strImportant & "</font>"
         '   NewField = Part1 & strImportantTags & Part2
          '  i = i + Len(strImportantTags) 
'        End If
 '       i = InStr(i + 1, NewField, "followed up", 1)
  '  Loop
    
    'tag "complete" - make green
'    i = InStr(1, NewField , "complete", 1)
'    Do While i <> 0
 '   	j = i
  '  	k = FindEnd(i + 1, NewField)
   '     If k > i + 1 Then
    '        Part1 = Left(NewField, j - 1)
     '       strImportant = Mid(NewField, j, k - j)
      '      Part2 = Mid(NewField,k)
       '     strImportantTags = "<font color='#008000'>" & strImportant & "</font>"
        '    NewField = Part1 & strImportantTags & Part2
         '   i = i + Len(strImportantTags)
'        End If
 '       i = InStr(i + 1, NewField, "complete", 1)
  '  Loop
    
    'tag "completed" - make green
'    i = InStr(1, NewField , "completed", 1)
 '   Do While i <> 0
  '  	j = i
   ' 	k = FindEnd(i + 1, NewField)
    '    If k > i + 1 Then
     '       Part1 = Left(NewField, j - 1)
      '      strImportant = Mid(NewField, j, k - j)
       '     Part2 = Mid(NewField,k)
        '    strImportantTags = "<font color='#008000'>" & strImportant & "</font>"
         '   NewField = Part1 & strImportantTags & Part2
          '  i = i + Len(strImportantTags)
'        End If
 '       i = InStr(i + 1, NewField, "completed", 1)
  '  Loop
    
    InsertTags = NewField

End Function

Function FindEnd(StartPos, StrText)
    FindEnd = Instr(StartPos + 1, StrText, " ")
    n = Instr(StartPos + 1, StrText, ".")
    If n < FindEnd And n <> 0 then FindEnd = n
    n = Instr(StartPos + 1, StrText, ")")
    If n < FindEnd And n <> 0 then FindEnd = n
End Function

Function FindEndIgnorePeriod(StartPos, StrText)
    FindEndIgnorePeriod = Instr(StartPos + 1, StrText, " ")
    n = Instr(StartPos + 1, StrText, ")")
    If n < FindEndIgnorePeriod And n <> 0 then FindEndIgnorePeriod = n
End Function

%>