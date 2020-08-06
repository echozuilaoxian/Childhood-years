<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head><title><% name=session("user") %>谢谢您！您已经成功退出了论坛！</title></head><meta http-equiv=refresh content=3;url=default.asp>
<body background="img/background.jpg">
<font size="2"> 
<% user=request.querystring("user") %>
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select* from online where user='"&user&"'"
rs.Open sql,conn,1,3
rs.delete
rs.close
set conn=nothing 
%>
</font> 
<div align="center"><font size="2" color="#FF0000"><%=session("user")%>谢谢您，您已经成功退出了论坛！</font><font size="2"> 
  <font color="#FF0000">本页将自动转向，请等待…………</font></font></div>
<% Session.Abandon
session("user")="" %></body>
</html>