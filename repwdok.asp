<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>找回你的密码第三步--论坛系统</title>
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
      <a href="default.asp">首页</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="reg.asp">注册</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>登陆</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">退出</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">搜索</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">留言</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">管理员入口</a> 
      </font></td>
  </tr>
</table>
<br>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center" height="23">
  <tr> 
    <td height="21"><b>&gt;&gt;欢迎光临本论坛</b></td>
  </tr>
  <tr> 
    <td height="27" bgcolor="#F5F5F5"><a href="default.asp">童真岁月论坛</a> → <a href="repwd.asp">找回个人密码步骤一</a> 
      → <font color="#000000">找回个人密码步骤二</font><font color="#FF0000"> → 找回个人密码步骤三</font></td>
  </tr>
</table>
<br>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="32" bgcolor="#F5F5F5" colspan="3"> 
      <div align="left"><img src="img/BMsg.gif" width="16" height="16" align="absmiddle">通过下面的方式您可以找回自己丢失的密码::::::::::::::::::: 
        <% id=trim(request.form("id")) %>
		<% tishidaan=trim(request.form("tishidaan")) %>
        <% if tishidaan="" then
	  response.write"<script>alert('对不起，请正确输入您的密码提示答案');history.back(-1);</script>"
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
response.write"<font color=red>对不起，你输入的密码提示答案不正确，请<a href=javascript:history.back(-1)>后退</a>重新输入</font>"
response.End()
else
%>
      </div>
    </td>
  </tr>
  <tr> 
    <td height="37" width="242">→第三步：您的密码已经找回，<font color="#FF0000">请牢记</font></td>
    <td height="37" width="164"><font color="#000000">你的用户名：</font><font color="#FF0000"><%=rs("user")%></font></td>
    <td height="37" width="353"><font color="#000000">您的密码：</font><font color="#FF0000"><%=rs("pwd")%></font></td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#F5F5F5" colspan="3">
      <div align="center"><font size="2" color="#000000">※</font><font color="#FF0000">恭喜您！您的密码已经找回，请务必牢记，很高兴为您服务<font size="2" color="#000000">※</font></font></div>
    </td>
  </tr>
</table>
<br>
<br>
<table width="763" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td bgcolor="#F5F5F5" height="24"> 
      <div align="center"><font color="#000000">| 网站设计 | 友情链接 | 给我留言 | QQ：349643649 
        |</font></div>
    </td>
  </tr>
  <tr> 
    <td height="18"> 
      <div align="center">&copy; 2006 Childhood Of Saint Bbs 童真岁月ASP官方论坛 版权所有</div>
    </td>
  </tr>
</table>
<% end if %><% rs.close %>
</body>
</html>
