<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<% if session("user")="" then
response.redirect"adminlogin.asp"
end if %><html>
<head>
<title>大区修改--童真岁月论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from lei where id="&request.querystring("id")
on error resume next 
rs.Open sql,conn,1
%>
<br>
<form name="form1" method="post" action="leimodifyok.asp">
  <table width="328" border="0" height="70">
    <tr> 
      <td width="95" bgcolor="#F5F5F5"> 
        <div align="center"><font size="2" color="#000000">大区编号：</font></div>
      </td>
      <td width="233"> <font size="2"> 
        <input type="text" name="id" value="<%=rs("id")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="95" bgcolor="#F5F5F5"> 
        <div align="center"><font size="2" color="#000000">大区名称：</font></div>
      </td>
      <td width="233"> <font size="2"> 
        <input type="text" name="typename" value="<%=rs("typename")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="95"><font size="2"></font></td>
      <td width="233"> <font size="2"> 
        <input type="submit" name="Submit" value="修改">
        <input type="reset" name="Submit2" value="Reset">
        </font></td>
    </tr>
  </table>
</form>
<%
rs.close
%></body>
</html>
