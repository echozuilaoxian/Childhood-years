<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>������Ϣ--��̳ϵͳ</title>
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
</STYLE></head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="10">
<table width="755" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td><img src="img/banner.jpg" width="760" height="80" border="1"></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="27" bgcolor="#CCCCCC"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%> 
      <font color="#000000"> </font></font><img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="default.asp">��ҳ</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="reg.asp">ע��</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>��½</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">�˳�</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">����Ա���</a></font></td>
  </tr>
</table>
<br>
<% user=request.querystring("user") %>
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select* from user where user='"&user&"'"
on error resume next
rs.open sql,conn,1
%>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="21" colspan="3"><b>&gt;&gt; ��ӭ���ٱ���̳</b></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#F5F5F5" colspan="3">ͯ��������̳ �� �����������</td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#FFFFFF" width="396" colspan="2"><img src="img/banzhu.gif" width="19" height="17" align="absmiddle">��<font color="#FF0000"><%=rs("user")%><font color="#000000">�ĸ�������</font></font></td>
    <td height="28" width="364">
      <div align="right"><font color="#000000"><img src="images/stats.gif" width="16" height="16" align="absmiddle">��ǰ״̬��</font><font color="#FF0000"> 
        <% rs("user")=user %>
        <% dim rs2
dim sql2
set rs2=server.createobject("adodb.recordset")
sql2="select* from online where user='"&user&"'"
on error resume next
rs2.open sql2,conn,1
if rs2.eof then
response.write"[����]"
else
response.write"[����]"
end if 
rs2.close %>
        </font> </div>
    </td>
  </tr>
</table>
<table width="760" border="1" align="center" bordercolor="#F5F5F5" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="28" bgcolor="#F5F5F5" colspan="2"><b>�������ϣ�</b></td>
    <td height="224" rowspan="7" width="356">&nbsp;</td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#FFFFFF" width="121"> 
      <div align="right"><font color="#000000">�ǳƣ�</font></div>
    </td>
    <td height="28" bgcolor="#FFFFFF" width="275"><font color="#FF0000"><%=rs("user")%></font></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#F5F5F5" width="121"> 
      <div align="right"><font color="#000000">�Ա�</font></div>
    </td>
    <td height="28" bgcolor="#F5F5F5" width="275"><font color="#FF0000"><%=rs("sex")%></font></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#FFFFFF" width="121"> 
      <div align="right"><font color="#000000">E-mail:</font></div>
    </td>
    <td height="28" bgcolor="#FFFFFF" width="275"><font color="#FF0000"><%=rs("email")%></font></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#F5F5F5" width="121"> 
      <div align="right"><font color="#000000">��ѸQQ��</font></div>
    </td>
    <td height="28" bgcolor="#F5F5F5" width="275"><font color="#FF0000"><%=rs("QQ")%></font></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#FFFFFF" width="121"> 
      <div align="right"><font color="#000000">������ҳ��</font></div>
    </td>
    <td height="28" bgcolor="#FFFFFF" width="275"><font color="#FF0000"><%=rs("homepage")%></font></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#F5F5F5" width="121"> 
      <div align="right"><font color="#000000">��ַ��</font></div>
    </td>
    <td height="28" bgcolor="#F5F5F5" width="275"><font color="#FF0000"><%=rs("address")%></font></td>
  </tr>
</table>
<br>
<table width="756" border="1" align="center" bordercolor="#F5F5F5" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25" bgcolor="#F5F5F5" colspan="4"><b>��̳���ԣ�</b></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#FFFFFF" width="119"> 
      <div align="right"><font color="#000000">����ͷ��</font></div>
    </td>
    <td height="50" bgcolor="#FFFFFF" rowspan="2" width="276"><font color="#FF0000"><img src=images/<%=rs("userico")%>.jpg border="1"></font></td>
    <td height="25" bgcolor="#FFFFFF" width="124"> 
      <div align="right"><font color="#000000">�ȼ���</font></div>
    </td>
    <td height="25" bgcolor="#FFFFFF" width="227"><font color="#FF0000">
      <% if rs("GDP")>200 then
response.write"�߼���ʦ"
end if
if rs("GDP")<200 and rs("GDP")>100 then
response.write"�м���ʦ"
end if
if rs("GDP")<100 then
response.write"ħ��ѧͽ"
end if %>
      </font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#F5F5F5" width="119"><font color="#000000"></font></td>
    <td height="25" bgcolor="#F5F5F5" width="124"> 
      <div align="right"><font color="#000000">ע��ʱ�䣺</font></div>
    </td>
    <td height="25" bgcolor="#F5F5F5" width="227"><font color="#FF0000"><%=rs("time")%></font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#FFFFFF" width="119"> 
      <div align="right"><font color="#000000">����ֵ��</font></div>
    </td>
    <td height="25" bgcolor="#FFFFFF" width="276"><font color="#FF0000"><%=rs("GDP")%></font></td>
    <td height="25" bgcolor="#FFFFFF" width="124"> 
      <div align="right"><font color="#000000">�������£�</font></div>
    </td>
    <td height="25" bgcolor="#FFFFFF" width="227"><font color="#FF0000"> 
      <% rs("user")=user %>
      <% dim rs3
dim sql3
set rs3=server.createobject("adodb.recordset")
sql3="select count(id) as sum from wenzhang where name='"&user&"'"
on error resume next
rs3.open sql3,conn,1
%>
      <%=rs3("sum")%> 
      <% rs3.close %>
      ƪ</font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#F5F5F5" width="119"> 
      <div align="right"><font color="#000000">�Ǽ���</font></div>
    </td>
    <td height="25" bgcolor="#F5F5F5" width="276"><font color="#000000"> 
      <% if rs("GDP")<100 then
		  response.write"<font color=red>��</font>������"
		  end if
		  if rs("GDP")<200 and rs("GDP")>100 then
		  response.write"<font color=red>���</font>�����"
		  end if
		  if rs("GDP")<300 and rs("GDP")>200 then
		  response.write"<font color=red>����</font>����"
		  end if
		  if rs("GDP")<400 and rs("GDP")>300 then
		  response.write"<font color=red>�����</font>���"
		  end if 
		  if rs("GDP")<500 and rs("GDP")>400 then
		  response.write"<font color=red>������</font>��"
		  end if
		  if rs("GDP")>500 then
		  response.write"<font color=red>�������</font>"
		  end if %>
      </font></td>
    <td height="25" bgcolor="#F5F5F5" width="124"> 
      <div align="right"><font color="#000000">�ظ����£�</font></div>
    </td>
    <td height="25" bgcolor="#F5F5F5" width="227"><font color="#FF0000">
      <% rs("user")=user %>
      <% dim rs4
dim sql4
set rs4=server.createobject("adodb.recordset")
sql4="select count(id) as sum from rwenzhang where rname='"&user&"'"
on error resume next
rs4.open sql4,conn,1
%>
      <%=rs4("sum")%> 
      <% rs4.close %>
      ƪ</font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#FFFFFF" width="119"> 
      <div align="right"><font color="#000000">����ǩ����</font></div>
    </td>
    <td height="25" bgcolor="#FFFFFF" colspan="3"><font size="2" color="#FF0000">��</font><font color="#FF0000"><%=rs("sign")%><font size="2">��</font></font></td>
  </tr>
</table>
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
<% rs.close %>
</body>
</html>
