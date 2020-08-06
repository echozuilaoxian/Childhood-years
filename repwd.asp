<%@ CODEPAGE ="936"%>
<html>
<head>
<title>找回密码功能--论坛系统</title>
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
    <td height="27" bgcolor="#F5F5F5"><a href="default.asp">童真岁月论坛</a> → <font color="#FF0000">找回个人密码步骤一</font></td>
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
    <td height="37" width="186">→第一步：请输入您的用户名：</td>
    <td height="37" width="573">&nbsp;</td>
  </tr>
  <tr> 
    <td height="39" colspan="2">
      <form name="form1" method="post" action="repwd2.asp">
        请输入用户名： 
        <input type="text" name="user" size="30" STYLE="background-color:transparent;border:1px solid #000000">
        <input type="submit" name="Submit" value=" 提交查询 " STYLE="background-color:transparent;border:1px solid #000000">
      </form>
    </td>
  </tr>
</table>
<br>
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
</body>
</html>
