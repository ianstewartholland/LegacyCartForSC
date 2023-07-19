<%
Function ReplaceSpaces(TextString)

    Dim StringInProcess
    
    StringInProcess = TextString

    'replace spaces
    SpacePos = Instr(1, StringInProcess, " ")
    Do While SpacePos > 0
        FirstPart = Left(StringInProcess, SpacePos - 1)
        SecondPart = Mid(StringInProcess, SpacePos + 1)
        StringInProcess = FirstPart & "%20" & SecondPart
        SpacePos = Instr(1, StringInProcess, " ")
    Loop
    'replace ampersands
    AmpPos = Instr(1, StringInProcess, "&")
    Do While AmpPos > 0
        FirstPart = Left(StringInProcess, AmpPos - 1)
        SecondPart = Mid(StringInProcess, AmpPos + 1)
        StringInProcess = FirstPart & "***" & SecondPart
        AmpPos = Instr(1, StringInProcess, "&")
    Loop
    
    ReplaceSpaces = StringInProcess
    
End Function
%>