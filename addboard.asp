<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<% if session("user")="" then
response.redirect"adminlogin.asp"
end if %>
<head>
<title>增加新看版--童真岁月论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="addboardok.asp">
  <table width="399" border="0" height="73">
    <tr> 
      <td width="92" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">看版名称：</font></div>
      </td>
      <td width="307"> <font size="2"> 
        <input type="text" name="boardname">
        </font></td>
    </tr>
    <tr> 
      <td height="21" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">图标：</font></div>
      </td>
      <td> 
        <input type="text" name="ico">
      </td>
    </tr>
    <tr> 
      <td height="21" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">斑竹：</font></div>
      </td>
      <td> 
        <input type="text" name="banzhu">
      </td>
    </tr>
    <tr> 
      <td height="21" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">看版简介：</font></div>
        <div align="right"></div>
      </td>
      <td> <font size="2"> 
        <textarea name="boardcontent" cols="40" rows="4"></textarea>
        </font></td>
    </tr>
    <tr> 
      <td width="92" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">所属大区：</font></div>
      </td>
      <td width="307"> <font size="2"> 
        <input type="text" name="typeid" size="5">
        <font color="#FF0000"> *(注意：请输入此版所属大区的编号)</font></font></td>
    </tr>
    <tr> 
      <td width="92"><font size="2"></font></td>
      <td width="307"> <font size="2"> 
        <input type="submit" name="Submit" value="增加">
        <input type="reset" name="Submit2" value="Reset">
        </font></td>
    </tr>
  </table>
</form>
<br>
<font size="2" color="#FF0000">论坛大区编号： 
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select* from lei"
on error resume next
rs.open sql,conn,1
do while not rs.eof
%>
<br>
</font>
<table width="401" border="0" cellspacing="0" cellpadding="0" height="10">
  <tr> 
    <td width="73" height="7">&nbsp;</td>
    <td width="328" height="7"><font size="2" color="#FF0000"><%=rs("id")%><%=rs("typename")%></font></td>
  </tr>
</table>
<font size="2" color="#FF0000"> </font>
<% 
rs.movenext
loop
rs.close
%>
<br>
</body>
</html>
