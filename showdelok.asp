<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--> 
<% id=request.form("id") %>
<html>
<head>
<title>您已经删除帖子成功，请稍后……</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=3;url=showadmin.asp?id=<%=request.form("boardid")%>>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<font color="#FF0000"> 
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select*from wenzhang where id="&id&""
rs.Open sql,conn,1,3 
rs.delete
rs.close
%>
<%
dim rs2 
dim sql2 
set rs2 = server.createobject("adodb.recordset") 
sql2 = "select*from rwenzhang where rid="&id&""
rs2.Open sql2,conn,1,3
if not rs2.eof then 
rs2.delete
rs2.close
end if
%>
<font size="2">您已经删除帖子成功，请稍后……</font></font> 
</body>
</html>

