<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<font size="2">ͯ��������̳&gt;&gt;&gt;<a href="addboard.asp">�����¿���</a></font><br>
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
      <div align="center"><font size="2" color="#000000">�������ƣ�</font></div>
    </td>
    <td width="400"><font size="2"><%=rs("boardname")%></font></td>
    <td rowspan="3"> 
      <div align="center"><font size="2">[<a href="boardmodify.asp?id=<%=rs("id")%>">�޸�</a>]<br>
        [<a href="boarddel.asp?id=<%=rs("id")%>">ɾ��</a>]</font></div>
      <font size="2"></font></td>
  </tr>
  <tr> 
    <td width="86" bgcolor="#F5F5F5"> 
      <div align="center"><font size="2" color="#000000">�����飺</font></div>
    </td>
    <td width="400"><font size="2"><%=rs("boardcontent")%></font></td>
  </tr>
  <tr> 
    <td width="86" bgcolor="#F5F5F5"> 
      <div align="center"><font size="2" color="#000000">�������</font></div>
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
