<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>滚动信息修改--论坛管理系统--学习论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=3;url=notemodify.asp>
</head>
<body bgcolor="#FFFFFF" text="#000000" background="img/background.jpg">
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select*from tongzhi"
rs.Open sql,conn,1,3 
rs("note")=request.form("note")
rs.update
rs.close
%>
<font size="2" color="#FF0000">谢谢使用！您已经修改成功！请稍后……</font> 
</body>
</html>

