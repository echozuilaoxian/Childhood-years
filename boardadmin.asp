<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<font size="2">童真岁月论坛&gt;&gt;&gt;<a href="addboard.asp">增加新看版</a></font><br>
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from board "
on error resume next 
rs.Open sql,conn,1
RS.MoveFirst
do while not rs.eof
%>
<table width="562" border="1" bordercolorlight="#FFFFFF" bordercolordark="#000000" cellspacing="0" cellpadding="0" bordercolor="#000000">
  <tr> 
    <td width="86" bgcolor="#F5F5F5"> 
      <div align="center"><font size="2" color="#000000">版面名称：</font></div>
    </td>
    <td width="400"><font size="2"><%=rs("boardname")%></font></td>
    <td rowspan="3"> 
      <div align="center"><font size="2">[<a href="boardmodify.asp?id=<%=rs("id")%>">修改</a>]<br>
        [<a href="boarddel.asp?id=<%=rs("id")%>">删除</a>]</font></div>
      <font size="2"></font></td>
  </tr>
  <tr> 
    <td width="86" bgcolor="#F5F5F5"> 
      <div align="center"><font size="2" color="#000000">看版简介：</font></div>
    </td>
    <td width="400"><font size="2"><%=rs("boardcontent")%></font></td>
  </tr>
  <tr> 
    <td width="86" bgcolor="#F5F5F5"> 
      <div align="center"><font size="2" color="#000000">本版斑竹：</font></div>
    </td>
    <td width="400"><font size="2"><%=rs("banzhu")%></font></td>
  </tr>
</table>
<p> 
  <% 
rs.movenext  
loop  
rs.Close  
%>
</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
