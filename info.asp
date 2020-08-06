<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>个人信息--论坛系统</title>
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
      <a href="default.asp">首页</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="reg.asp">注册</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>登陆</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">退出</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">搜索</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">留言</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">管理员入口</a></font></td>
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
    <td height="21" colspan="3"><b>&gt;&gt; 欢迎光临本论坛</b></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#F5F5F5" colspan="3">童真岁月论坛 → 浏览个人资料</td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#FFFFFF" width="396" colspan="2"><img src="img/banzhu.gif" width="19" height="17" align="absmiddle">：<font color="#FF0000"><%=rs("user")%><font color="#000000">的个人资料</font></font></td>
    <td height="28" width="364">
      <div align="right"><font color="#000000"><img src="images/stats.gif" width="16" height="16" align="absmiddle">当前状态：</font><font color="#FF0000"> 
        <% rs("user")=user %>
        <% dim rs2
dim sql2
set rs2=server.createobject("adodb.recordset")
sql2="select* from online where user='"&user&"'"
on error resume next
rs2.open sql2,conn,1
if rs2.eof then
response.write"[离线]"
else
response.write"[在线]"
end if 
rs2.close %>
        </font> </div>
    </td>
  </tr>
</table>
<table width="760" border="1" align="center" bordercolor="#F5F5F5" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="28" bgcolor="#F5F5F5" colspan="2"><b>基本资料：</b></td>
    <td height="224" rowspan="7" width="356">&nbsp;</td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#FFFFFF" width="121"> 
      <div align="right"><font color="#000000">昵称：</font></div>
    </td>
    <td height="28" bgcolor="#FFFFFF" width="275"><font color="#FF0000"><%=rs("user")%></font></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#F5F5F5" width="121"> 
      <div align="right"><font color="#000000">性别：</font></div>
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
      <div align="right"><font color="#000000">腾迅QQ：</font></div>
    </td>
    <td height="28" bgcolor="#F5F5F5" width="275"><font color="#FF0000"><%=rs("QQ")%></font></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#FFFFFF" width="121"> 
      <div align="right"><font color="#000000">个人主页：</font></div>
    </td>
    <td height="28" bgcolor="#FFFFFF" width="275"><font color="#FF0000"><%=rs("homepage")%></font></td>
  </tr>
  <tr> 
    <td height="28" bgcolor="#F5F5F5" width="121"> 
      <div align="right"><font color="#000000">地址：</font></div>
    </td>
    <td height="28" bgcolor="#F5F5F5" width="275"><font color="#FF0000"><%=rs("address")%></font></td>
  </tr>
</table>
<br>
<table width="756" border="1" align="center" bordercolor="#F5F5F5" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25" bgcolor="#F5F5F5" colspan="4"><b>论坛属性：</b></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#FFFFFF" width="119"> 
      <div align="right"><font color="#000000">个性头像：</font></div>
    </td>
    <td height="50" bgcolor="#FFFFFF" rowspan="2" width="276"><font color="#FF0000"><img src=images/<%=rs("userico")%>.jpg border="1"></font></td>
    <td height="25" bgcolor="#FFFFFF" width="124"> 
      <div align="right"><font color="#000000">等级：</font></div>
    </td>
    <td height="25" bgcolor="#FFFFFF" width="227"><font color="#FF0000">
      <% if rs("GDP")>200 then
response.write"高级大法师"
end if
if rs("GDP")<200 and rs("GDP")>100 then
response.write"中级法师"
end if
if rs("GDP")<100 then
response.write"魔法学徒"
end if %>
      </font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#F5F5F5" width="119"><font color="#000000"></font></td>
    <td height="25" bgcolor="#F5F5F5" width="124"> 
      <div align="right"><font color="#000000">注册时间：</font></div>
    </td>
    <td height="25" bgcolor="#F5F5F5" width="227"><font color="#FF0000"><%=rs("time")%></font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#FFFFFF" width="119"> 
      <div align="right"><font color="#000000">经验值：</font></div>
    </td>
    <td height="25" bgcolor="#FFFFFF" width="276"><font color="#FF0000"><%=rs("GDP")%></font></td>
    <td height="25" bgcolor="#FFFFFF" width="124"> 
      <div align="right"><font color="#000000">主题文章：</font></div>
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
      篇</font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#F5F5F5" width="119"> 
      <div align="right"><font color="#000000">星级：</font></div>
    </td>
    <td height="25" bgcolor="#F5F5F5" width="276"><font color="#000000"> 
      <% if rs("GDP")<100 then
		  response.write"<font color=red>☆</font>☆☆☆☆☆"
		  end if
		  if rs("GDP")<200 and rs("GDP")>100 then
		  response.write"<font color=red>☆☆</font>☆☆☆☆"
		  end if
		  if rs("GDP")<300 and rs("GDP")>200 then
		  response.write"<font color=red>☆☆☆</font>☆☆☆"
		  end if
		  if rs("GDP")<400 and rs("GDP")>300 then
		  response.write"<font color=red>☆☆☆☆</font>☆☆"
		  end if 
		  if rs("GDP")<500 and rs("GDP")>400 then
		  response.write"<font color=red>☆☆☆☆☆</font>☆"
		  end if
		  if rs("GDP")>500 then
		  response.write"<font color=red>☆☆☆☆☆☆</font>"
		  end if %>
      </font></td>
    <td height="25" bgcolor="#F5F5F5" width="124"> 
      <div align="right"><font color="#000000">回复文章：</font></div>
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
      篇</font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#FFFFFF" width="119"> 
      <div align="right"><font color="#000000">个性签名：</font></div>
    </td>
    <td height="25" bgcolor="#FFFFFF" colspan="3"><font size="2" color="#FF0000">※</font><font color="#FF0000"><%=rs("sign")%><font size="2">※</font></font></td>
  </tr>
</table>
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
<% rs.close %>
</body>
</html>
