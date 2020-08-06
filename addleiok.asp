<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>大区添加成功，请稍后…………</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=2;url=adminlei.asp></head>

<body bgcolor="#FFFFFF" text="#000000">
<% leiname=Request.Form("leiname") %>
<% if leiname="" then
response.write"<script>alert('对不起，请填写完整');history.back(-1);</script>"
response.End()
else %>
<% set save=conn.execute("insert into lei(typename)values('"&leiname&"')")%>
<% set save=nothing %>
<% end if %>
<font size="2" color="#FF0000">大区添加成功！请稍后………… </font> 
</body>
</html>
