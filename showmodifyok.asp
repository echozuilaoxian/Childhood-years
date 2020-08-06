<!--#include file="conn.asp"-->
<% id=request.form("id") %><html>
<head>
<title>修改成功--童真岁月论坛！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=3;url=showadmin.asp?id=<%=request.form("boardid")%>>
</head>

<body bgcolor="#FFFFFF" text="#000000" background="img/background.jpg">
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select*from wenzhang where id="&id&""
rs.Open sql,conn,1,3
rs("boardid")=request.form("boardid") 
rs("name")=request.form("name")
rs("title")=request.form("title")
rs("content")=request.form("content")
rs("top")=request.form("top")
rs("lock")=request.form("lock")
rs.update
rs.close
%>
<table width="527" border="0" align="center" height="29">
  <tr>
    <td><font size="2" color="#FF0000">您已经修改成功！谢谢使用！本页面将自动转向，请等待……</font></td>
  </tr>
</table>
</body>
</html>
