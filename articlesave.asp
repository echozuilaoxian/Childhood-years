<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>лл���������޸��Ѿ��ɹ������Ժ󡭡�</title>
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
��<a href="javascript:history.go(-1)">����</a>����д�������ϣ������ܷ�������! 
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
<% response.write"<font color=red>���Ѿ��޸ĳɹ������Ժ󡭡�</font>" %>
<% end if %>
<% set conn=nothing %></body>
</html>
