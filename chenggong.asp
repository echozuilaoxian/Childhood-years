<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>处理修改--请稍后…………</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=3;url=boardadmin.asp></head>

<body bgcolor="#FFFFFF" text="#000000">
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from board where id="&request.form("id")
rs.Open sql,conn,1,3
rs("boardname")=request.form("boardname")
rs("ico")=request.form("ico")
rs("banzhu")=request.form("banzhu")
rs("boardcontent")=request.form("boardcontent")
rs("typeid")=request.form("typeid")
rs("lock")=request.form("lock")
rs.update
rs.close
%>
<font size="2" color="#FF0000">修改成功！请稍后………… </font> 
</body>
</html>
