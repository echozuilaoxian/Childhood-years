
<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<if session("user")="" then
response.redirect"adminlogin.asp"
end if %>
<html>
<head>
<title>滚动信息管理--论坛管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<br>
<% dim rs 
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from tongzhi"
rs.open sql,conn,1
rs.movefirst
%>
<table width="561" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="54" bgcolor="#F5F5F5"><font size="2">童真岁月论坛&gt;&gt;&gt;论坛管理系统&gt;&gt;&gt;滚动信息管理&gt;&gt;&gt;</font></td>
  </tr>
  <tr>
    <td>
      <form name="form1" method="post" action="notemodifyok.asp">
        <table width="546" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>
              <input type="text" name="note" size="60" value="<%=rs("note")%>" maxlength="200">
              <input type="submit" name="Submit" value="修改">
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<% rs.close %></body>
</html>

