<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<% if session("user")="" then
response.redirect"adminlogin.asp"
end if %><html>
<head>
<title>删除大区--童真岁月论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from lei where id="&request.querystring("id")
rs.Open sql,conn,1
%>
<form name="form1" method="post" action="leidelok.asp">
  <table width="400" border="0" height="56">
    <tr> 
      <td width="102" bgcolor="#F5F5F5"> 
        <div align="center"><font size="2" color="#000000">大区编号：</font></div>
      </td>
      <td width="298"> <font size="2"> 
        <input type="text" name="id" value="<%=rs("id")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="102" bgcolor="#F5F5F5"> 
        <div align="center"><font size="2" color="#000000">大区名称：</font></div>
      </td>
      <td width="298"> <font size="2"> 
        <input type="text" name="typename" value="<%=rs("typename")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="102"><font size="2"></font></td>
      <td width="298"> <font size="2"> 
        <input type="submit" name="Submit" value="删除">
        <input type="reset" name="Submit2" value="Reset">
        </font></td>
    </tr>
  </table>
</form>
<% 
rs.close
%></body>
</html>
