'--------------------------------------
' SEND URL TO HTTP SERVER
'--------------------------------------

Function SendHTTP(sRequest)

	Set xmlHttp = WScript.CreateObject("MSXML2.ServerXMLHTTP")

	xmlHttp.Open "GET", sRequest, False
	xmlHttp.Send ""
	getHTML = xmlHttp.responseText
	status = xmlHttp.status
	xmlHttp.Abort

	Set xmlHttp = Nothing
	
	If status = 200 Then
		If Len(getHTML) > 0 Then
			SendHTTP = getHTML
		else 
			SendHTTP = "Erreur serveur, retour vide"
		End If
	else
			SendHTTP = "Erreur serveur, status " & status
	End If


End Function
