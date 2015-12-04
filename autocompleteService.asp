<%
Dim strInput: strInput = Request("input")
Dim JSON: JSON = Request("json")
Dim id: id = ""
Dim region: region = ""
Dim location: location = ""
Dim states: states = ""
Dim site: site = ""
Set objXML = Server.CreateObject("Microsoft.XMLDOM")
Set objRows = Server.CreateObject("Microsoft.XMLDOM")
Set objSite = Server.CreateObject("Microsoft.XMLDOM")

objXML.async = False
objXML.Load (Server.MapPath("worldheritage.xml"))

If objXML.parseError.errorCode <> 0 Then
	Response.Write error
End If

Set objRows = objXML.getElementsByTagName("row")

noOfRows = objRows.length

If (JSON) Then
	Response.ContentType="text/html"
	Response.Write "{ results: ["
	For i = 0 To (noOfRows - 1)
		id = Server.HTMLEncode(objRows.item(i).getAttribute("id"))
		region = Server.HTMLEncode(objRows.item(i).childNodes(0).text)
		location = Server.HTMLEncode(objRows.item(i).childNodes(1).text)
		states = Server.HTMLEncode(objRows.item(i).childNodes(2).text)
		site = Server.HTMLEncode(objRows.item(i).childNodes(3).text)
		info = region&" - "&states
		If(InStr(LCase(site),LCase(strInput))) Then
			Response.Write "{ id: """&id&""", value: """&site&""", regn: """&info&""" },"
		End If
	Next
	Response.Write "] }"			
Else
	Response.ContentType="text/xml"
	Response.Write "<results>"
	For i = 0 To (noOfRows - 1)
		id = objRows.item(i).getAttribute("id")
		region = Server.HTMLEncode(objRows.item(i).childNodes(0).text)
		location = Server.HTMLEncode(objRows.item(i).childNodes(1).text)
		states = Server.HTMLEncode(objRows.item(i).childNodes(2).text)
		site = Server.HTMLEncode(objRows.item(i).childNodes(3).text)
		locstate = location & ", " & state
		If(InStr(LCase(site),LCase(strInput))) Then
			Response.Write "<rs id="""&id&""" locstate="""&region&""" regn="""&region&""">"&site&"</rs>"
		End If
	Next
	Response.Write "</results>"
End If
%>
