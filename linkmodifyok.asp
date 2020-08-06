<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>修改成功，请稍后…………</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=3;url=linkadmin.asp>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select*from link where id="&request.form("id")
rs.Open sql,conn,1,3
rs("sitename")=request.form("sitename")
rs("siteurl")=request.form("siteurl")
rs("logourl")=request.form("logourl")
rs.update
rs.close
%>
<font size="2" color="#FF0000">您已经修改成功，请稍后…………</font> 
</body>
</html>
