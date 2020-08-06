<!--#include file="conn.asp"-->
<html>
<head>
<title>作者相关文章检索--论坛系统</title>
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
    <td height="28" bgcolor="#F5F5F5" colspan="3">童真岁月论坛 → 作者相关文章检索</td>
  </tr>
</table>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="25"><img src="img/banzhu.gif" width="19" height="17" align="absmiddle">：<font color="#FF0000"><%=rs("user")%><font color="#000000">发表的所有文章，本次搜索共找到</font> 
      <% user=request.querystring("user") %>
      <% dim rs2
dim sql2
set rs2=server.createobject("adodb.recordset")
sql2="select count(id) as sum from wenzhang where name='"&user&"'"
on error resume next
rs2.open sql2,conn,1
%>
      <%=rs2("sum")%> 
      <% rs2.close %>
      <font color="#000000">条主题，</font> 
      <% user=request.querystring("user") %>
      <% dim rs3
dim sql3
set rs3=server.createobject("adodb.recordset")
sql3="select count(id) as sum from rwenzhang where rname='"&user&"'"
on error resume next
rs3.open sql3,conn,1
%>
      <%=rs3("sum")%> 
      <% rs3.close %>
      <font color="#000000">条回复</font></font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#F5F5F5"><font color="#FF0000"><%=rs("user")%></font>相关检索：主题列表<font color="#FF0000"> 
      <% user=request.querystring("user") %>
      <%
dim rs4
dim sql4
set rs4 = server.createobject("adodb.recordset")
sql4 = "select*from wenzhang where bid=0 and name='"&user&"' order by id desc"
count=conn.execute("select count(id) from wenzhang where bid=0 and name='"&user&"'")(0)
on error resume next
pagesetup=10
rs4.Open sql4,conn,1
if rs4.eof then
response.write"对不起，没有相关主题"
end if
If Count/pagesetup > (Count\pagesetup) then
TotalPage=(Count\pagesetup)+1
else TotalPage=(Count\pagesetup)
End If
PageCount= 0
RS4.MoveFirst
if Request.QueryString("ToPage")<>"" then PageCount = cint(Request.QueryString("ToPage"))
if PageCount <=0 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
RS4.Move (PageCount-1) * pagesetup
i=1
do while not rs4.eof
%>
      </font></td>
  </tr>
</table>
<table width="759" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="25" bgcolor="#FFFFFF"><img src="img/BMsg.gif" width="16" height="16"> 
      <font color="#FF0000" size="3"><a href="show.asp?id=<%=rs4("id")%>"><%=rs4("title")%></a> </font></td>
  </tr>
  <tr>
    <td height="25" bgcolor="#F5F5F5">　 作者：<font color="#FF0000"><%=rs4("name")%>→『 
      <% board=rs4("boardid")%>
      <% dim rsb
	  dim sqlb
	  set rsb=server.createobject("adodb.recordset")
	  sqlb="select* from board where id="&board&""
	  on error resume next
	  rsb.open sqlb,conn,1
	  %>
      <%=rsb("boardname")%> 
      <% rsb.close %>
      』 发表于：<%=rs4("time")%></font></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#FFFFFF"> 
      <div align="center"><font color="#CCCCCC">-----------------------------------------------------------------------------------------------------------------------------</font></div>
    </td>
  </tr>
</table>
<%i=i+1 
if i>pagesetup then exit do 
rs4.movenext 
loop 
rs4.Close 
%>
<br>
<table width="755" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>分页：
      <% user=request.querystring("user")%>
	  <% 
ii=PageCount-5 
iii=PageCount+5 
if ii < 1 then 
ii=1 
end if 
if iii > TotalPage then 
iii=TotalPage 
end if 
if PageCount > 6 then 
Response.Write "<a href=?user="&user&"&topage=1><font color=black>1</font></a> ... " 
end if 
 
for i=ii to iii 
If i<>PageCount then 
Response.Write "<a href=?user="&user&"&topage="& i &"><font color=black>" & i & "</font></a> " 
else 
Response.Write " <font color=red><b>"&i&"</b></font> " 
end if 
next 
 
if TotalPage > PageCount+5 then 
Response.Write " ... <a href=?user="&user&"&topage="&TotalPage&"><font color=black>"&TotalPage&"</font></a>" 
end if 
%>
    </td>
  </tr>
</table>
<br>
<% rs.close %>
<br>
<table width="755" border="0" cellspacing="0" cellpadding="0" align="center">
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

