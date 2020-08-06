<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>谢谢您，您的修改已经成功，请稍后……</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=3;url=show.asp?id=<%=request.form("id")%>>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<% name=Replace(Request.Form("name"),"'","''")
title=Replace(Request.Form("title"),"'","''")
biaoqing=Replace(Request.Form("biaoqing"),"'","''")
content=Replace(Request.Form("GuestContent"),"'","''")
id=Replace(Request.Form("id"),"'","''")
%>
<% if name="" or title="" or content="" then %>
请<a href="javascript:history.go(-1)">后退</a>请填写完整资料，您才能发表文章! 
<%else%>
<BR><% id=request.form("id") %>
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select* from wenzhang where id="&id&""
rs.open sql,conn,1,3
rs("title")=Replace(Request.Form("title"),"'","''")
rs("biaoqing")=Replace(Request.Form("biaoqing"),"'","''")
rs("content")=Replace(Request.Form("GuestContent"),"'","''")
rs.update
rs.close
%>
<% response.write"<font color=red>您已经修改成功！请稍后……</font>" %>
<% end if %>
<% set conn=nothing %></body>
</html>
