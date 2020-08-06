<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<% if session("user")="" then
response.redirect"adminlogin.asp"
end if %>
<head>
<title>看版删除--童真岁月论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select* from board where id="&request.querystring("id")
on error resume next
rs.open sql,conn,1
do while not rs.eof
%>
<form name="form1" method="post" action="boarddelok.asp">
  <table width="392" border="0" height="87">
    <tr> 
      <td width="98" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">看版编号：</font></div>
      </td>
      <td width="294"> 
        <input type="text" name="id" value="<%=rs("id")%>">
      </td>
    </tr>
    <tr> 
      <td width="98" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">看版名称：</font></div>
      </td>
      <td width="294"> <font size="2"> 
        <input type="text" name="boardname" value="<%=rs("boardname")%>">
        </font></td>
    </tr>
    <tr> 
      <td bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">看版简介：</font></div>
        <div align="right"></div>
      </td>
      <td> <font size="2"> 
        <textarea name="boardcontent" cols="30" rows="4"><%=rs("boardcontent")%></textarea>
        </font></td>
    </tr>
    <tr> 
      <td width="98" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">所属大区：</font></div>
      </td>
      <td width="294"> <font size="2"> 
        <input type="text" name="typeid" size="5" value="<%=rs("typeid")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="98"><font size="2" color="#000000"></font></td>
      <td width="294"> <font size="2"> 
        <input type="submit" name="Submit" value="删除">
        <input type="reset" name="Submit2" value="Reset">
        </font></td>
    </tr>
  </table>
</form>
<% 
rs.movenext
loop
rs.close
%></body>
</html>

