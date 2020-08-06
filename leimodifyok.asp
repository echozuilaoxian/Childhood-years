<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>大区修改成功！请稍后…………</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=2;url=adminlei.asp></head>

<body bgcolor="#FFFFFF" text="#000000">
<% if request.form("typename")="" then
response.write"<script>alert('对不起，请填写完整');history.back(-1);</script>"
response.End()
end if %>
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from lei where id="&request.form("id")
on error resume next 
rs.Open sql,conn,1,3
rs("typename")=request.form("typename")
rs.update
rs.close
%>
<font size="2" color="#FF0000">大区修改成功！请稍后………… </font> 
</body>
</html>
