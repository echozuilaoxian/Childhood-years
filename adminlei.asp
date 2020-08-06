<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<% if session("user")="" then
response.redirect"adminlogin.asp"
end if %>
<html>
<head>
<title>大区管理--童真岁月论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<font size="2">学习论坛&gt;&gt;&gt;<a href="addlei.asp">增加新大区</a></font><br>
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from lei "
on error resume next 
rs.Open sql,conn,1
RS.MoveFirst
do while not rs.eof
%>
<br>
<table width="422" border="1" cellspacing="0" cellpadding="0" height="34">
  <tr> 
    <td width="93" bgcolor="#F5F5F5"> 
      <div align="center"><font size="2" color="#000000">大区编号：</font></div>
    </td>
    <td width="222"><font size="2"><%=rs("id")%></font></td>
    <td rowspan="2" width="107"> 
      <div align="center"><font size="2">[<a href="leimodify.asp?id=<%=rs("id")%>">修改</a>][<a href="leidel.asp?id=<%=rs("id")%>">删除</a>]</font></div>
    </td>
  </tr>
  <tr> 
    <td width="93" bgcolor="#F5F5F5"> 
      <div align="center"><font size="2" color="#000000">大区名称：</font></div>
    </td>
    <td width="222"><font size="2"><%=rs("typename")%></font></td>
  </tr>
</table>
<% 
rs.movenext
loop
rs.close
%></body>
</html>

