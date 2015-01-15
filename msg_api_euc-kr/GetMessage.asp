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
data.add "msg_serial", "씨리얼 키"
data.add "list_count", "가져올 갯수"
data.add "page", "페이지 번호"

Set messageService = New MessageService
messageService.getToken client_id, api_key

result = messageService.getMessage(data)

Response.Write result

%>