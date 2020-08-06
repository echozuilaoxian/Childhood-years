<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--> 
<%
dim rs4 
dim sql4 
set rs4 = server.createobject("adodb.recordset") 
sql4 = "select*from board where id="&Trim(Request.QueryString("id")) 
on error resume next 
rs4.Open sql4,conn,1 
RS4.MoveFirst 
do while not rs4.eof 
%>
<html>
<head>
<title><%=rs4("boardname")%>专区--童真岁月论坛！</title>
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
    <td height="33" bgcolor="#CCCCCC" background="images/tbg.gif"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%> 
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
      <% user=session("user") %>
      <% if session("user")="" then response.write"<a href=login.asp>登陆</a>" else  response.write"<a href=out.asp?user="&user&">退出</a>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">搜索</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="http://my.phome.cn/knight999/guest/index.asp" target="_blank">留言</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">管理员入口</a></font></td>
  </tr>
</table>
<font size="2" color="#FFFFFF"> </font> <br>
<table width="762" border="0" cellspacing="0" cellpadding="0" align="center" height="23">
  <tr> 
    <td height="2"> 
      <form name="form1" method="post" action="loginok.asp">
        <table width="761" border="0" height="19" style="border: 1pt solid #333333" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="29" bgcolor="#CCCCCC" width="761" background="images/tbg.gif"><img src="img/nofollow.gif" width="15" height="15">快速登陆入口：</td>
          </tr>
          <tr> 
            <td height="36" bgcolor="#FFFFFF" width="761"> 
              <div align="left"><font color="#000000"> <img src="img/banzhu.gif" width="19" height="17"> 
                用户名： 
                <input type="text" name="user" size="10" style="background-color:transparent;border:1px solid #000000">
                用户： 
                <input type="text" name="pwd" size="10" style="background-color:transparent;border:1px solid #000000">
                <input type="submit" name="Submit" value="登陆" style="background-color:transparent;border:1px solid #000000">
                [<a href="reg.asp">没有注册</a>][<a href="repwd.asp">忘记密码</a>]</font></div>
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<table width="762" border="0" cellspacing="0" cellpadding="0" align="center" height="15">
  <tr> 
    <td width="456"> 
      <% dim rs9 
	   dim sql9
	   set rs9=server.createobject("adodb.recordset")
	   sql9="select*from tongzhi"
	   rs9.open sql9,conn,1
	   %>
      <font color="#0080ff"> </font><marquee hspace=5 loop=100 scrollamount=2 scrolldelay=5 width=90% on="left" align="center" DIRECTI><font color="#000000"> 
      </font><font size="2" color="#666666"><%=rs9("note")%></font><font color="#0080ff"> 
      <%
	set co5=server.createobject("adodb.recordset")
	co5.open "select count(id) as sum from user",conn,1,3
	if not co5.eof then
%>
      </font> </marquee></td>
    <td width="306"> 
      <div align="right"><font color="#000000">论坛共有注册会员：</font><font color="#3399CC"><%=co5("sum")%><font color="#000000">位</font> 
        <%
	end if
	co5.close
	set co5=nothing
	%>
        <% dim rs0 
	   dim sql0
	   set rs0=server.createobject("adodb.recordset")
	   sql0="select*from user"
	   rs0.open sql0,conn,1
	   rs0.movelast
	   %>
        <font color="#000000">，欢迎新会员</font><%=rs0("user")%><font color="#000000">加入</font></font></div>
    </td>
  </tr>
</table>
<table width="763" border="0" bordercolorlight="#FFFFFF" bordercolordark="#3399CC" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td height="28" colspan="2"><font color="#000000"><img src="img/nofollow.gif" width="15" height="15"><a href="default.asp">童真岁月论坛</a> 
      &gt;&gt;&gt;<%=rs4("boardname")%>&gt;&gt;&gt;<a href="write.asp?id=<%=rs4("id")%>"><img src="img/newanc.gif" width="60" height="17" align="absmiddle" border="0"></a></font> 
  <tr > 
    <td width="25" height="24" bgcolor="#F5F5F5" background="images/tbg.gif"> 
      <div align="center"><font color="#FFFFFF">状态</font></div>
    </td>
    <td width="307" height="24" background="images/tbg.gif">
      <div align="center"><font size="2" color="#FFFFFF">贴子标题</font></div>
    </td>
    <td width="80" height="24" bgcolor="#FFFFFF" background="images/tbg.gif"> 
      <div align="center"><font size="2" color="#FFFFFF">作者</font></div>
    </td>
    <td width="36" height="24" bgcolor="#FFFFFF" background="images/tbg.gif"> 
      <div align="center"><font size="2" color="#FFFFFF">点击</font></div>
    </td>
    <td width="37" height="24" bgcolor="#FFFFFF" background="images/tbg.gif"> 
      <div align="center"><font size="2" color="#FFFFFF">回复</font></div>
    </td>
    <td width="251" height="24" bgcolor="#FFFFFF" background="images/tbg.gif"> 
      <div align="center"><font size="2" color="#FFFFFF">发表时间</font><font color="#FFFFFF"><font size="2"> 
        </font></font></div>
    </td>
  </tr>
</table>
<font color="#0080ff"><font size="2" color="#000000"> 
<% idd=request.querystring("id")%>
</font><font color="#0080ff"><font size="2"> 
<%
dim rs8 
dim sql8 
set rs8 = server.createobject("adodb.recordset") 
sql8 = "select*from wenzhang where top=1 and boardid="&idd&" order by id desc"
on error resume next 
rs8.Open sql8,conn,1 
RS8.MoveFirst 
do while not rs8.eof 
%>
</font></font><font size="2" color="#000000"> </font><font color="#000000"> 
<% rid=rs8("id")%>
<%
	set co4=server.createobject("adodb.recordset")
	co4.open "select count(rid) as sum from rwenzhang where rid="&rid&"",conn,1,3
	if not co4.eof then
%>
</font></font> 
<table width="762" style="word-break:break-all" border="0" bordercolorlight="#3399CC" bordercolordark="#003399" align="center" bordercolor="#99CC99" cellpadding="0" cellspacing="1" bgcolor="#333333">
  <tr> 
    <td width="25" height="15" bgcolor="#FFFFFF"> 
      <div align="center"><img src="img/parttop.gif" width="18" height="18"></div>
    </td>
    <td width="307" height="15" bgcolor="#FFFFFF"><font size="2"><img src="img/<%=rs8("biaoqing")%>.gif"><a href=show.asp?id=<%=rs8("id")%>&boardid=<%=rs4("id")%>><%=left(rs8("title"),24)%></a></font></td>
    <td width="80" height="15" bgcolor="#FFFFFF"> 
      <div align="center"><font size="2"><%=rs8("name")%></font></div>
    </td>
    <td width="36" height="15" bgcolor="#FFFFFF"> 
      <div align="center"><font size="2"><%=rs8("hits")%></font></div>
    </td>
    <td width="37" height="15" bgcolor="#FFFFFF"> 
      <div align="center"><font size="2" color="#000000"><%=co4("sum")%></font></div>
    </td>
    <td width="251" height="15" bgcolor="#FFFFFF"> 
      <div align="left"><font size="2"><%=rs8("time")%></font><font size="2" color="#000000"> 
        <%
	end if
	co4.close
	set co4=nothing
	%>
        最后回复： 
        <% iddd=rs8("id") %>
        <% dim rsh 
		dim sqlh
		set rsh=server.createobject("adodb.recordset")
		sqlh="select * from rwenzhang where rid="&iddd&""
		rsh.open sqlh,conn,1
		on error resume next
		if rsh.eof then
		response.write"------"
		end if
		rsh.movelast 
		%>
        <%=rsh("rname")%> 
        <% 
		rsh.close %>
        </font></div>
    </td>
  </tr>
</table>
<% 
rs8.movenext  
loop  
rs8.Close  
%>
<font size="2" color="#000000">
<% idd=request.querystring("id")%>
<%
dim rs
dim sql
set rs = server.createobject("adodb.recordset")
sql = "select*from wenzhang where bid=0 and boardid="&idd&" and top=0 order by id desc"
count=conn.execute("select count(id) from wenzhang where bid=0 and boardid="&idd&" and top=0")(0)
on error resume next
pagesetup=30
rs.Open sql,conn,1
If Count/pagesetup > (Count\pagesetup) then
TotalPage=(Count\pagesetup)+1
else TotalPage=(Count\pagesetup)
End If
PageCount= 0
RS.MoveFirst
if Request.QueryString("ToPage")<>"" then PageCount = cint(Request.QueryString("ToPage"))
if PageCount <=0 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
RS.Move (PageCount-1) * pagesetup
i=1
do while not rs.eof
%>
</font><font color="#000000"> 
<% rid=rs("id")%>
<%
	set co2=server.createobject("adodb.recordset")
	co2.open "select count(rid) as sum from rwenzhang where rid="&rid&"",conn,1,3
	if not co2.eof then
%>
</font> 
<div align="center"><font size="2" color="#000000"> </font> </div>
<table width="762" style="word-break:break-all" border="0" bordercolorlight="#3399CC" bordercolordark="#003399" align="center" bordercolor="#99CC99" cellspacing="1" cellpadding="0" bgcolor="#333333">
  <tr> 
    <td width="25" height="15" bgcolor="#FFFFFF"> 
      <div align="center"><img src="img/tpc.gif" width="18" height="18"></div>
    </td>
    <td width="307" height="15" bgcolor="#FFFFFF"><img src="img/<%=rs("biaoqing")%>.gif"><font size="2"><a href=show.asp?id=<%=rs("id")%>&boardid=<%=rs4("id")%>><%=left(rs("title"),24)%></a></font></td>
    <td width="80" height="15" bgcolor="#FFFFFF"> 
      <div align="center"><font size="2"><%=rs("name")%></font></div>
    </td>
    <td width="36" height="15" bgcolor="#FFFFFF"> 
      <div align="center"><font size="2"><%=rs("hits")%></font></div>
    </td>
    <td width="37" height="15" bgcolor="#FFFFFF"> 
      <div align="center"><font size="2" color="#000000"><%=co2("sum")%></font></div>
    </td>
    <td width="251" height="15" bgcolor="#FFFFFF"> 
      <div align="left"><font size="2"><%=rs("time")%>最后回复： 
        <% idddd=rs("id") %>
        <% dim rsg 
		dim sqlg
		set rsg=server.createobject("adodb.recordset")
		sqlg="select * from rwenzhang where rid="&idddd&""
		rsg.open sqlg,conn,1
		on error resume next
		if rsg.eof then
		response.write"-------"
		end if
		rsg.movelast 
		%>
        <%=rsg("rname")%> 
        <% 
		rsg.close %>
        </font></div>
    </td>
  </tr>
</table>
<font size="2" color="#000000"> 
<%
	end if
	co2.close
	set co2=nothing
	%>
</font> 
<%i=i+1 
if i>pagesetup then exit do 
rs.movenext 
loop 
rs.Close 
%>
<table width="200" border="0" cellspacing="0" cellpadding="0" align="right">
  <tr> 
    <td height="20"> 分页: 
      <% idd=request.querystring("id")%>
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
Response.Write "<a href=?id="&idd&"&topage=1><font color=black>1</font></a> ... " 
end if 
 
for i=ii to iii 
If i<>PageCount then 
Response.Write "<a href=?id="&idd&"&topage="& i &"><font color=black>" & i & "</font></a> " 
else 
Response.Write " <font color=red><b>"&i&"</b></font> " 
end if 
next 
 
if TotalPage > PageCount+5 then 
Response.Write " ... <a href=?id="&idd&"&topage="&TotalPage&"><font color=black>"&TotalPage&"</font></a>" 
end if 
%>
    </td>
  </tr>
</table>
<% 
rs4.movenext  
loop  
rs4.Close  
%>
<br>
<br>
<font color="#FFFFFF"> </font><br>
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


