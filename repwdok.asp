<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>�һ�������������--��̳ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 12pt
}
A:link {
	COLOR: #003366; TEXT-DECORATION: none
}
A:visited {
	COLOR: #003366; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ee9c00; TEXT-DECORATION: underline
}
</STYLE>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="10">
<table width="755" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td><img src="img/banner.jpg" width="760" height="80" border="1"></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="24" bgcolor="#FFFFFF"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%> 
      <% 
dim rsu
dim sqlu
set rsu=server.createobject("adodb.recordset")
sqlu="select* from online where user='"&session("user")&"'"
rsu.Open sqlu,conn,1,3
if not rsu.eof then
rsu("ltime")=now()
rsu.update
end if
rsu.close
%>
      <font color="#000000"> </font></font><img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="default.asp">��ҳ</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="reg.asp">ע��</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>��½</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">�˳�</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">����Ա���</a> 
      </font></td>
  </tr>
</table>
<br>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center" height="23">
  <tr> 
    <td height="21"><b>&gt;&gt;��ӭ���ٱ���̳</b></td>
  </tr>
  <tr> 
    <td height="27" bgcolor="#F5F5F5"><a href="default.asp">ͯ��������̳</a> �� <a href="repwd.asp">�һظ������벽��һ</a> 
      �� <font color="#000000">�һظ������벽���</font><font color="#FF0000"> �� �һظ������벽����</font></td>
  </tr>
</table>
<br>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="32" bgcolor="#F5F5F5" colspan="3"> 
      <div align="left"><img src="img/BMsg.gif" width="16" height="16" align="absmiddle">ͨ������ķ�ʽ�������һ��Լ���ʧ������::::::::::::::::::: 
        <% id=trim(request.form("id")) %>
		<% tishidaan=trim(request.form("tishidaan")) %>
        <% if tishidaan="" then
	  response.write"<script>alert('�Բ�������ȷ��������������ʾ��');history.back(-1);</script>"
	  response.End()
	  end if %>
        <% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select* from user where id="&id&"" 
on error resume next
rs.open sql,conn,1
%>
<% tishidaan=trim(request.form("tishidaan")) %>
<% if tishidaan<>rs("tishidaan") then
response.write"<font color=red>�Բ����������������ʾ�𰸲���ȷ����<a href=javascript:history.back(-1)>����</a>��������</font>"
response.End()
else
%>
      </div>
    </td>
  </tr>
  <tr> 
    <td height="37" width="242">�������������������Ѿ��һأ�<font color="#FF0000">���μ�</font></td>
    <td height="37" width="164"><font color="#000000">����û�����</font><font color="#FF0000"><%=rs("user")%></font></td>
    <td height="37" width="353"><font color="#000000">�������룺</font><font color="#FF0000"><%=rs("pwd")%></font></td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#F5F5F5" colspan="3">
      <div align="center"><font size="2" color="#000000">��</font><font color="#FF0000">��ϲ�������������Ѿ��һأ�������μǣ��ܸ���Ϊ������<font size="2" color="#000000">��</font></font></div>
    </td>
  </tr>
</table>
<br>
<br>
<table width="763" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td bgcolor="#F5F5F5" height="24"> 
      <div align="center"><font color="#000000">| ��վ��� | �������� | �������� | QQ��349643649 
        |</font></div>
    </td>
  </tr>
  <tr> 
    <td height="18"> 
      <div align="center">&copy; 2006 Childhood Of Saint Bbs ͯ������ASP�ٷ���̳ ��Ȩ����</div>
    </td>
  </tr>
</table>
<% end if %><% rs.close %>
</body>
</html>
