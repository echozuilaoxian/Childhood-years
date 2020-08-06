<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--> 
<html>
<head>
<title>搜索结果---童真岁月论坛！</title>
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
<body bgcolor="#FFFFFF" text="#000000">
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
<table width="763" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" align="center" bordercolor="#003399">
  <tr> 
    <td bgcolor="#F5F5F5" height="24"><font color="#000000"><img src="img/nofollow.gif" width="15" height="15"><font size="2">童真岁月论坛&gt;&gt;&gt;内容检索&gt;&gt;&gt;搜索结果</font></font></td>
  </tr>
</table>
<% word=Replace(Request.Form("word"),"'","''")
searchtype=request.form("searchtype") %>
<% if word="" then
response.write"<script>alert('对不起，请输入查询关键字');history.back(-1);</script>"
response.End()
end if %>
<br>
<table width="761" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" align="center">
  <tr> 
    <td width="316" height="21" bgcolor="#F5F5F5"> 
      <div align="center"><font color="#000000">标题</font></div>
    </td>
    <td width="121" height="21" bgcolor="#F5F5F5"> 
      <div align="center"><font color="#000000">作者</font></div>
    </td>
    <td width="74" height="21" bgcolor="#F5F5F5"> 
      <div align="center"><font color="#000000">点击</font></div>
    </td>
    <td width="240" height="21" bgcolor="#F5F5F5"> 
      <div align="center"><font color="#000000">发表时间</font></div>
    </td>
  </tr>
</table>
<% if searchtype="title" then
dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from wenzhang where title like '%"&word&"%'"
on error resume next 
rs.Open sql,conn,1
if rs.eof then
response.write"<font size=2 color=#FF0000>对不起，没有找到相关资料</font>"
end if 
RS.MoveFirst
do while not rs.eof
%>
<table width="759" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" align="center">
  <tr> 
    <td width="315" bgcolor="#F5F5F5"> <img src="img/Forum_readme.gif" width="10" height="10"><a href="show.asp?id=<%=rs("id")%>"><%=rs("title")%></a> 
    </td>
    <td width="123" bgcolor="#F5F5F5"> 
      <div align="center"><%=rs("name")%></div>
    </td>
    <td width="73" bgcolor="#F5F5F5"> 
      <div align="center"><%=rs("hits")%></div>
    </td>
    <td width="238" bgcolor="#F5F5F5"> 
      <div align="center"><%=rs("time")%> </div>
    </td>
  </tr>
</table>
<% 
rs.movenext  
loop  
rs.Close  
end if %>
<% if searchtype="content" then
dim rs2
dim sql2
set rs2=server.createobject("adodb.recordset")
sql2="select*from wenzhang where content like '%"&word&"%'"
on error resume next 
rs2.Open sql2,conn,1
if rs2.eof then
response.write"<font size=2 color=#FF0000>对不起，没有找到相关资料</font>"
end if 
RS2.MoveFirst 
do while not rs2.eof 
%>
<table width="759" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" align="center">
  <tr> 
    <td width="314" bgcolor="#F5F5F5"> <img src="img/Forum_readme.gif" width="10" height="10"><a href="show.asp?id=<%=rs2("id")%>"><%=rs2("title")%></a> 
    </td>
    <td width="125" bgcolor="#F5F5F5"> 
      <div align="center"><%=rs2("name")%></div>
    </td>
    <td width="72" bgcolor="#F5F5F5"> 
      <div align="center"><%=rs2("hits")%></div>
    </td>
    <td width="238" bgcolor="#F5F5F5"> 
      <div align="center"><%=rs2("time")%></div>
    </td>
  </tr>
</table>
<% 
rs2.movenext  
loop  
rs2.Close  
end if %>
<p>
</p>
<p> <br>
</p>
<table width="761" border="0" cellspacing="0" cellpadding="0" align="center">
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
<p>&nbsp; </p>
</body>
</html>

