<!--#include file="conn.asp"--><html>
<head>
<title>��̳��������--ͯ��������̳</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#CCCCCC" text="#000000" topmargin="0">
<table width="161" border="0" cellspacing="0" cellpadding="0" height="139">
  <tr> 
    <td height="19"> 
      <div align="center"><font size="2" color="#000000">::��̳��������::</font></div>
    </td>
  </tr>
  <tr> 
    <td height="20"> 
      <div align="center"><font size="2" color="#000000">��<a href="adminout.asp" target="_parent">�˳�����</a>��</font></div>
    </td>
  </tr>
  <tr> 
    <td height="24"> 
      <div align="center"><font size="2" color="#000000">��<a href="boardadmin.asp" target="mainFrame">�������</a>����<a href="adminlei.asp" target="mainFrame">��������</a>��</font></div>
    </td>
  </tr>
  <tr> 
    <td height="24"> 
      <div align="center"><font size="2" color="#000000">��<a href="linkadmin.asp" target="mainFrame">�������ӹ���</a>��</font></div>
    </td>
  </tr>
  <tr> 
    <td height="24"> 
      <div align="center"><font size="2" color="#000000">��<a href="notemodify.asp" target="mainFrame">�������Ź���</a>�� 
        <% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from board "
on error resume next 
rs.Open sql,conn,1
RS.MoveFirst
do while not rs.eof
%>
        </font></div>
    </td>
  </tr>
</table>
<table width="162" border="0" cellspacing="0" cellpadding="0" height="18">
  <tr>
    <td height="18"> 
      <div align="center"><font size="2" color="#000000">��<a href="showadmin.asp?id=<%=rs("id")%>" target="mainFrame"><%=rs("boardname")%></a>��</font></div>
    </td>
  </tr>
</table>
<% 
rs.movenext  
loop  
rs.Close  
%>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
