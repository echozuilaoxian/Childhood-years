<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"--><html>
<head><meta http-equiv=refresh content=2;url=login.asp>
<title>лл�������Ѿ�ע��ɹ���--�ʽ�������������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 12pt
}
A:link {
	COLOR: #0080ff; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0080ff; TEXT-DECORATION: none
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
    <td height="27" bgcolor="#FFFFFF"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%><font color="#000000"> 
      </font></font><img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
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
<table width="775" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="24" bgcolor="#F5F5F5"><font color="#000000"><img src="img/nofollow.gif" width="15" height="15"><font size="2">ͯ��������̳&gt;&gt;&gt;�û�ע��&gt;&gt;&gt;ע��ɹ�</font></font></td>
  </tr>
</table>
<table width="775" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><% 
user=Replace(trim(Request.Form("user")),"'","''")
pwd=md5(request.form("pwd"))
sex=Replace(Request.Form("sex"),"'","''")
QQ=Replace(Request.Form("QQ"),"'","''")
email=Replace(Request.Form("email"),"'","''")
mimatishi=Replace(Request.Form("mimatishi"),"'","''")
tishidaan=Replace(Request.Form("tishidaan"),"'","''")
homepage=Replace(Request.Form("homepage"),"'","''")
address=Replace(Request.Form("address"),"'","''")
userico=Replace(Request.Form("userico"),"'","''")
sign=Replace(Request.Form("sign"),"'","''")
%>
</td>
  </tr>
</table>
<p><% if user="" or pwd="" or QQ="" or email="" or mimatishi="" or tishidaan="" or userico="" then
response.write"<script language=javascript>alert('����д����');history.back(-1);</script>"
response.End()
else
user=Replace(trim(Request.Form("user")),"'","''")
pwd=md5(request.form("pwd"))
sex=Replace(Request.Form("sex"),"'","''")
QQ=Replace(Request.Form("QQ"),"'","''")
email=Replace(Request.Form("email"),"'","''")
mimatishi=Replace(Request.Form("mimatishi"),"'","''")
tishidaan=Replace(Request.Form("tishidaan"),"'","''")
homepage=Replace(Request.Form("homepage"),"'","''")
address=Replace(Request.Form("address"),"'","''")
userico=Replace(Request.Form("userico"),"'","''")
sign=Replace(Request.Form("sign"),"'","''")
dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from user where user='"&user&"'"
rs.open sql,conn,1
if not rs.eof then
response.write"<script>alert('�������Բ����û����Ѿ���ռ�ã���ʹ�������û���');history.back(-1);</script>"
response.End()
%>
<% 
end if
rs.close %>
<% end if %>
<% set reg=conn.execute("insert into user(user,pwd,sex,QQ,email,homepage,userico,address,mimatishi,tishidaan,sign)values('"&user&"','"&pwd&"','"&sex&"','"&QQ&"','"&email&"','"&homepage&"','"&userico&"','"&address&"','"&mimatishi&"','"&tishidaan&"','"&sign&"')")
response.write"<font size=2>���Ѿ�ע��ɹ������Ժ󡭡�</font>"
response.End()
%>
<% set reg=nothing %></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>
  
  <br>
</p>
<table width="775" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td bgcolor="#F5F5F5" height="24"> 
      <div align="center"><font color="#000000">| ��վ��� | �������� | �������� | QQ��47598622 
        |</font></div>
    </td>
  </tr>
  <tr> 
    <td height="18"> 
      <div align="center">&copy; 2006 Childhood Of Saint Bbs ͯ������ASP�ٷ���̳ ��Ȩ����</div>
    </td>
  </tr>
</table>
</body>
</html>
