<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--> <html>
<head>
<title>看版删除成功！请稍后…………</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=2;url=boardadmin.asp></head>

<body bgcolor="#FFFFFF" text="#000000">
<font color="#FF0000">
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from board where id="&request.form("id") 
rs.Open sql,conn,1,3
rs.delete
rs.close
%>
<font size="2">看版删除成功！请稍后…………</font></font> 
</body>
</html>
