<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>请回答您的密码提示答案--论坛找回密码系统</title>
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
    <td height="27" bgcolor="#FFFFFF"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%> 
      <font color="#000000"> </font></font><img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="default.asp">首页</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="reg.asp">注册</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>登陆</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">退出</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">搜索</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">留言</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">管理员入口</a></font></td>
  </tr>
</table>
<br>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center" height="23">
  <tr> 
    <td height="21"><b>&gt;&gt;欢迎光临本论坛</b></td>
  </tr>
  <tr> 
    <td height="27" bgcolor="#F5F5F5"><a href="default.asp">童真岁月论坛</a> → <a href="repwd.asp">找回个人密码步骤一</a> 
      → <font color="#FF0000">找回个人密码步骤二</font></td>
  </tr>
</table>
<br>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="32" bgcolor="#F5F5F5" colspan="2"> 
      <div align="left"><img src="img/BMsg.gif" width="16" height="16" align="absmiddle">通过下面的方式您可以找回自己丢失的密码:::::::::::::::::::</div>
    </td>
  </tr>
  <tr> 
    <td height="37" width="242">→第二步：请正确输入您的密码提示答案</td>
    <td height="37" width="517"> 
      <% user=trim(request.form("user")) %>
	  <% if user="" then
	  response.write"<script>alert('对不起，请输入您的用户名');history.back(-1);</script>"
	  response.End()
	  end if %>
      <% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select* from user where user='"&user&"'"
on error resume next
rs.open sql,conn,1
if rs.eof then
response.write"<font color=red>对不起，没有查询到相关用户，请正确输入或者您还未注册</font>"
response.End()
end if
%>
    </td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#F5F5F5"> 
      <div align="right"><font color="#FF0000">[</font><font color="#FF0000"><%=rs("user")%>]<font color="#000000">您的密码提示问题是：</font></font></div>
    </td>
    <td height="22" bgcolor="#F5F5F5"><font color="#FF0000"><%=rs("mimatishi")%></font></td>
  </tr>
  <tr> 
    <td height="41"> 
      <div align="right"><font color="#000000">请正确输入您的密码提示答案：</font></div>
    </td>
    <td height="41"> 
      <form name="form1" method="post" action="repwdok.asp">
        <br>
        <input type="text" name="tishidaan" size="30" STYLE="background-color:transparent;border:1px solid #000000">
        <input type="submit" name="Submit" value=" 找回我的密码 " STYLE="background-color:transparent;border:1px solid #000000">
        <input type="hidden" name="id" value="<%=rs("id")%>">
      </form>
      <font color="#FF0000"></font></td>
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
<% rs.close %></body>
</html>
