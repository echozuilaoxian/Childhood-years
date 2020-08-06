<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"--><html>
<head><meta http-equiv=refresh content=2;url=login.asp>
<title>谢谢您！你已经注册成功！--问剑江湖武侠网！</title>
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
<table width="775" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="24" bgcolor="#F5F5F5"><font color="#000000"><img src="img/nofollow.gif" width="15" height="15"><font size="2">童真岁月论坛&gt;&gt;&gt;用户注册&gt;&gt;&gt;注册成功</font></font></td>
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
response.write"<script language=javascript>alert('请填写完整');history.back(-1);</script>"
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
response.write"<script>alert('！！！对不起，用户名已经被占用，请使用其他用户名');history.back(-1);</script>"
response.End()
%>
<% 
end if
rs.close %>
<% end if %>
<% set reg=conn.execute("insert into user(user,pwd,sex,QQ,email,homepage,userico,address,mimatishi,tishidaan,sign)values('"&user&"','"&pwd&"','"&sex&"','"&QQ&"','"&email&"','"&homepage&"','"&userico&"','"&address&"','"&mimatishi&"','"&tishidaan&"','"&sign&"')")
response.write"<font size=2>您已经注册成功，请稍后……</font>"
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
