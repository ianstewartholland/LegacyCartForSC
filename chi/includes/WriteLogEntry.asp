<%
'
' Write a log message to the logfile for debugging 
'
Function WriteLogEntry(strMessage)
    Dim logfile: logfile = "C:\WWWUsers\LocalUser\chi\computationalhydraulics_com\databases\commit.log"
    Dim f, fs
    
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(logfile, 8, true)
    
    f.WriteLine(Date & " " & Time & ": " & strMessage)
    f.Close()
    Set f = Nothing
    Set fs = Nothing
    WriteLogEntry = True
End Function
 %>