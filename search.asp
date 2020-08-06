<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>Welcome to 论坛搜索--童真岁月论坛系统</title>
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
<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<table width="755" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td><img src="img/banner.jpg" width="760" height="80" border="1"></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="27" bgcolor="#FFFFFF"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%><% 
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
%> <font color="#000000"> 
      </font></font><img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
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
<table width="761" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" align="center" bordercolor="#003399">
  <tr> 
    <td height="24" bgcolor="#F5F5F5"><font color="#000000"><img src="img/nofollow.gif" width="15" height="15"><font size="2">童真岁月论坛&gt;&gt;&gt;内容检索&gt;&gt;&gt;</font></font></td>
  </tr>
</table>
<br>
<form name="form1" method="post" action="result.asp">
  <table width="411" border="1" cellspacing="0" cellpadding="0" height="75" align="center" bordercolor="#F5F5F5">
    <tr bgcolor="#F5F5F5" > 
      <td colspan="2" height="19"> 
        <div align="center"><font color="#000000">:::::::::;:::::论坛内容--数据检索::::::::::::::</font></div>
      </td>
    </tr>
    <tr> 
      <td width="103" height="30"> 
        <div align="right"><font color="#000000"><img src="img/pic161.gif" width="32" height="32">检索入口：</font></div>
      </td>
      <td width="302" height="30"> 
        <div align="left"> 　 
          <input type="text" name="word">
          <select name="searchtype">
            <option  value="title">标题</option>
            <option  value="content">内容</option>
          </select>
          <input type="submit" name="Submit" value="搜索">
        </div>
      </td>
    </tr>
  </table>
</form>
<p><br>
</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="769" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="24" bgcolor="#F5F5F5"> 
      <div align="center"><font color="#000000">| 网站设计 | 友情链接 | 给我留言 | QQ：47598622 
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
