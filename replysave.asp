<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>лл���������޸��Ѿ��ɹ������Ժ󡭡�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<% name=Replace(Request.Form("name"),"'","''")
pwd=Replace(Request.Form("pwd"),"'","''")
content=Replace(Request.Form("GuestContent"),"'","''")
id=Replace(Request.Form("id"),"'","''")
%>
<% if name="" or pwd="" or content="" then %>
��<a href="javascript:history.go(-1)">����</a>����д�������ϣ������ܷ�������! 
<%else%>
<BR><% id=request.form("id") %>
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select* from rwenzhang where id="&id&""
rs.open sql,conn,1,3
rs("rcontent")=Replace(Request.Form("GuestContent"),"'","''")
rs.update
rs.close
%>
<% response.write"<font color=red>���Ѿ��޸ĳɹ������Ժ󡭡�</font>" %>
<% dim rs2
dim sql2
set rs2=server.createobject("adodb.recordset")
sql2="select* from rwenzhang where id="&request.form("id")
rs2.open sql2,conn,1
%>
<meta http-equiv=refresh content=3;url=show.asp?id=<%=rs2("rid")%>>
<% rs2.close %>
<% end if %>
<% set conn=nothing %></body>
</html>
