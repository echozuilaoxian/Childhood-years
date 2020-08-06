<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<% if session("user")="" then
response.redirect"adminlogin.asp"
end if %><head>
<title>友情连接管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<font size="2">童真岁月论坛&gt;&gt;&gt;后台管理&gt;&gt;&gt;友情链接管理&gt;&gt;&gt;<a href="morelink.asp">增加新链接</a>&gt;&gt;&gt;</font> 
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select*from link"
on error resume next 
rs.Open sql,conn,1 
RS.MoveFirst 
do while not rs.eof 
%>
<br>
<table width="421" border="1" bordercolorlight="#FFFFFF" bordercolordark="#000000" cellspacing="0" cellpadding="0" height="21">
  <tr> 
    <td height="17" width="89" bgcolor="#F5F5F5"><font size="2" color="#000000">合作站点：</font></td>
    <td height="17" width="240"><font size="2"><%=rs("sitename")%></font></td>
    <td height="34" rowspan="2"> 
      <div align="center"><font size="2">[<a href="linkmodify.asp?id=<%=rs("id")%>">修改</a>]</font></div>
    </td>
  </tr>
  <tr> 
    <td height="17" width="89" bgcolor="#F5F5F5"><font size="2" color="#000000">站点地址：</font></td>
    <td height="17"><font size="2"><%=rs("siteurl")%></font></td>
  </tr>
</table>
<% 
rs.movenext
loop
rs.close %>
<br>
</body>
</html>
