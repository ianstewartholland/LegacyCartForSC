<%
'Dim xmlDoc
'Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'xmlDoc.async = False
'xmlDoc.loadXML("<GetList><a><b attr=""5"">a node</b></a></GetList>")

'DisplayNode xmlDoc.childNodes


' =================== Helper Routines ==================

Public Function DisplayNode(Nodes)
  DisplayNode = DisplayNode_(Nodes, 0)
End Function

Function DisplayNode_(Nodes, Indent)
   Dim strOutput
   Dim xNode
   strOutput = ""
   For Each xNode In Nodes
      Select Case xNode.nodeType
        Case 1:   ' NODE_ELEMENT
          If xNode.nodeName <> "#document" Then
            ' change DisplayAttrs_(xNode, Indent + 2) to 
            ' DisplayAttrs_(xNode, 0) for inline attributes
            strOutput = strOutput & "<pre>" & strDup(" ", Indent)  & "&lt;" & xNode.nodeName & _
              DisplayAttrs_(xNode, Indent + 2) & "&gt;" & vbCrLf
            If xNode.hasChildNodes Then
              strOutput = strOutput & DisplayNode_(xNode.childNodes, Indent + 2)
            End If
            strOutput = strOutput & vbCrLf & strDup(" ", Indent) & "&lt;/" & xNode.nodeName & "&gt;" & "</pre>"
          Else 
            If xNode.hasChildNodes Then
              strOutput = strOutput & DisplayNode_(xNode.childNodes, Indent + 2)
            End If
          End If
        Case 3:   ' NODE_TEXT                       
          strOutput = strOutput & strDup(" ", Indent) & xNode.nodeValue
      End Select
   Next
   DisplayNode_ = strOutput
End Function

Function DisplayAttrs_(Node, Indent)
   Dim xAttr, res

   res = ""
   For Each xAttr In Node.attributes
      If Indent = 0 Then
        res = res & " " & xAttr.name & "=""" & xAttr.value & """"
      Else 
        res = res & vbCrLf & strDup(" ", Indent) & xAttr.name & _
          "=""" & xAttr.value & """"
      End If
   Next

   DisplayAttrs_ = res
End Function

Function strDup(dup, c)
  Dim res, i

  res = ""
  For i = 1 To c
    res = res & dup
  Next
  strDup = res
End Function
%>