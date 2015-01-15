<% @CODEPAGE="65001" language="vbscript" %>
<% Option Explicit %>
<% session.CodePage = "65001" %>
<% Response.CharSet = "utf-8" %>
<% Response.buffer = True %>
<% Response.Expires = 0 %>

<!-- #include file="lib/MessageService.asp" -->
<!-- #include file="common.asp" -->

<%

Dim messageService, result, data

Set data = Server.CreateObject("Scripting.Dictionary")
data.add "msg_type", "mms"
data.add "callback", ""
data.add "subject", ""
data.add "msg", ""
data.add "image", "D:\mms\sample.jpg"

'data.add "trandate", "20150101000000" '예약 전송

data.add "phone", "수신번호_1" '한 명 전송
'data.add "phone", "수신번호_1, 수신번호_2" '여러 명 전송

Set messageService = New MessageService
messageService.getToken client_id, api_key

result = ms.sendMessage(data)

Set data = Nothing

Response.Write result

%>