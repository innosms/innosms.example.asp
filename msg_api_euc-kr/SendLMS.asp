<% @CODEPAGE="949" language="vbscript" %>
<% Option Explicit %>
<% session.CodePage = "949" %>
<% Response.CharSet = "euc-kr" %>
<% Response.buffer = True %>
<% Response.Expires = 0 %>

<!-- #include file="lib/MessageService.asp" -->
<!-- #include file="common.asp" -->

<%

Dim messageService, result, data

Set data = Server.CreateObject("Scripting.Dictionary")
data.add "msg_type", "lms"
data.add "callback", ""
data.add "subject", ""
data.add "msg", ""

data.add "phone", "수신번호_1" '한 명 전송
'data.add "phone", "수신번호_1, 수신번호_2" '여러 명 전송

'data.add "trandate", "20150101000000" '예약 전송

Set messageService = New MessageService
messageService.getToken client_id, api_key

result = messageService.sendMessage(data)

Set data = Nothing

Response.Write result

%>