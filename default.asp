<%@ CODEPAGE ="936"%>
<%
dim t1,t2 
t1=now() %>
<!--#include file="conn.asp"-->
<html>
<head>
<title>Welcome to 童真岁月论坛！朋友，是否还记得小时候那首熟悉的童谣？是否还记得小时候美好快乐的时光？</title>
<SCRIPT LANUGAGE="JavaScript">
<!--
function pop(newpage) {
var
popwin=window.open(newpage,"popWin","scrollbars=no,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,width=500,height=115");
return false;
}
//-->
</SCRIPT><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 12pt
}
A:link {
	COLOR: #000000; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #666600; TEXT-DECORATION: underline
}
</STYLE>
</head>

<body text="#000000" topmargin="10" link="#FFFFFF">
<table width="755" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td><img src="img/banner.jpg" width="760" height="80" border="1"></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="32" bgcolor="#CCCCCC"><font size="2" color="#000000"><font color="#FFFFFF"><%=session("user")%> 
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
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="default.asp">首页</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="reg.asp">注册</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% user=session("user") %>
      <% if session("user")="" then response.write"<a href=login.asp>登陆</a>" else  response.write"<a href=out.asp?user="&user&">退出</a>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">搜索</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="http://my.phome.cn/knight999/guest/index.asp" target="_blank">留言</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">管理员入口</a></font></font></td>
  </tr>
</table>
<br>
<table width="763" border="0" cellspacing="0" cellpadding="0" align="center" height="23">
  <tr> 
    <td height="2">
      <form name="form1" method="post" action="loginok.asp">
        <table width="761" border="0" height="19" style="border: 1pt solid #333333" cellpadding="0" cellspacing="0">
          <tr>
            <td height="29" bgcolor="#CCCCCC" width="761" background="images/tbg.gif"><font color="#FFFFFF"><b><img src="img/nofollow.gif" width="15" height="15">快速登陆入口</b></font>：</td>
          </tr>
          <tr> 
            <td height="36" bgcolor="#FFFFFF" width="761"> 
              <div align="left"><font color="#000000"> <img src="img/banzhu.gif" width="19" height="17"> 
                用户名： 
                <input type="text" name="user" size="10" STYLE="background-color:transparent;border:1px solid #000000">
                用户： 
                <input type="text" name="pwd" size="10" STYLE="background-color:transparent;border:1px solid #000000">
                <input type="submit" name="Submit" value="登陆" STYLE="background-color:transparent;border:1px solid #000000">
                [<a href="reg.asp">没有注册</a>][<a href="repwd.asp">忘记密码</a>]</font></div>
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<table width="761" border="0" cellspacing="0" cellpadding="0" align="center" height="15">
  <tr> 
    <td width="456" bgcolor="#FFFFFF"> <font color="#0080ff"><img src="img/gonggao.gif" width="18" height="18"></font>
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
    <td width="305" bgcolor="#FFFFFF"> 
      <div align="right"><font color="#000000">论坛共有注册会员：</font><font color="#0066FF"><%=co5("sum")%><font color="#000000">位</font> 
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
<% rs0.close %>
<font color="#0080ff"> 
<% rs9.close %>
</font> 
<table width="761" border="0" style="border: 1pt solid #333333" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="29" bgcolor="#CCCCCC" background="images/tbg.gif"><font size="2" color="#FFFFFF"><img src="img/BMsg.gif" width="16" height="16"><b>::最新帖子&gt;&gt;&gt;</b></font></td>
    <td height="29" bgcolor="#CCCCCC" background="images/tbg.gif"><font size="2" color="#FFFFFF"><img src="img/BMsg.gif" width="16" height="16"><b>::精华贴子&gt;&gt;&gt;</b></font></td>
  </tr>
  <tr> 
    <td height="58" bgcolor="#FFFFFF"> <font size="2" color="#000000"> 
      <% dim rsx 
	  dim sqlx
set rsx=server.createobject("adodb.recordset")
sqlx="select top 10 * from wenzhang order by id desc"
rsx.open sqlx,conn,1
do while not rsx.eof
%>
      </font> 
      <table width="394" border="0" height="17" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="21"><font size="2" color="#000000"><img src="img/nofollow.gif" width="15" height="15"><a href="show.asp?id=<%=rsx("id")%>&boardid=<%=rsx("boardid")%>"><%=server.htmlencode(rsx("title"))%></a></font></td>
        </tr>
      </table>
      <font size="2" color="#000000"> 
      <% rsx.movenext
loop
rsx.close
%>
      </font></td>
    <td height="58" bgcolor="#FFFFFF"> <font size="2" color="#000000"> 
      <% dim rsj 
	  dim sqlj
set rsj=server.createobject("adodb.recordset")
sqlj="select top 10 * from wenzhang where top=1 order by id desc"
rsj.open sqlj,conn,1
do while not rsj.eof
%>
      </font> 
      <table width="368" border="0" height="21" cellspacing="0" cellpadding="0">
        <tr> 
          <td><font size="2" color="#000000"><img src="img/nofollow.gif" width="15" height="15"><a href="show.asp?id=<%=rsj("id")%>&boardid=<%=rsj("boardid")%>"><%=server.htmlencode(rsj("title"))%></a></font></td>
        </tr>
      </table>
      <font size="2" color="#000000"> 
      <% rsj.movenext
loop
rsj.close
%>
      </font></td>
  </tr>
</table>
<font color="#0080ff"><br>
</font> <font color="#000000"><font size="2"><font color="#FFFFFF"> 
<% dim rst       
dim sqlt
set rst=server.createobject("adodb.recordset")
sqlt="select* from lei"
on error resume next
rst.open sqlt,conn,1
do while not rst.eof
%>
</font></font></font><font color="#FFFFFF"><font size="2"> </font></font> 
<table width="764" border="0" style="border: 1pt solid #333333;border-top-width:0px" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" height="19">
  <tr> 
    <td bgcolor="#666666" height="30" width="528" background="images/tbg.gif"><font size="2" color="#FFFFFF"><b>[<img src="img/nofollow.gif" width="15" height="15">]==<%=rst("typename")%>==<%=rst("id")%></b></font></td>
    <td height="30" width="235" bgcolor="#666666" background="images/tbg.gif">&nbsp;</td>
  </tr>
</table>
<% idd=rst("id") %>
<font color="#FFFFFF"><font size="2"> 
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select*from board where typeid="&idd&""
on error resume next 
rs.Open sql,conn,1 
RS.MoveFirst 
do while not rs.eof 
%>
</font></font> 
<table width="764" border="0"  style="border: 1pt solid #333333;border-top-width:0px" cellspacing="0" align="center" bordercolor="#F5F5F5" height="31" cellpadding="0">
  <tr bgcolor="#FFFFFF" bordercolor="#F5F5F5"> 
    <td width="18" height="39" bgcolor="#F5F5F5"> 
      <div align="center"><font color="#000000">::<br>
        ::<br>
        ::<br>
        ::<br>
        ::</font></div>
    </td>
    <td height="39" width="52" bgcolor="#F5F5F5"> 
      <div align="center"><img src="img/forum.gif" width="29" height="29"></div>
    </td>
    <td height="39" width="451"> 
      <table width="447" style="word-break:break-all" border="0" height="91" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="19" width="382"><font size="2" color="#000000">『<a href=board.asp?id=<%=rs("id")%>><%=rs("boardname")%></a>』</font></td>
          <td height="50" rowspan="2" width="55"> 
            <div align="right"><font color="#000000"><img src="img/<%=rs("ico")%>" border="1" width="53" height="63"></font></div>
          </td>
        </tr>
        <tr> 
          <td height="18" style="word-break:break-all" width="382"><font size="2" color="#000000"><img src="img/Forum_readme.gif" width="10" height="10"><%=rs("boardcontent")%></font></td>
        </tr>
      </table>
    </td>
    <td height="39" width="242"> <font color="#000000"> 
      <% banhao=rs("id") %>
      <%
dim rs4
dim sql4
set rs4=server.createobject("adodb.recordset")
sql4="select*from wenzhang where boardid="&banhao&" order by time desc"
on error resume next
rs4.open sql4,conn,1
%>
      </font> 
      <table width="238" style="word-break:break-all" border="0" cellspacing="0" cellpadding="0" height="36">
        <tr> 
          <td height="22"><font color="#000000"><img src="img/elist.gif" width="12" height="13">最后话题： 
            <%=server.htmlencode(rs4("title"))%></font></td>
        </tr>
        <tr> 
          <td height="16"><font color="#000000"><img src="img/post.gif" width="12" height="13">作者：<%=rs4("name")%></font></td>
        </tr>
        <tr> 
          <td height="5"><font color="#000000"><img src="img/post.gif" width="12" height="13">时间：<%=rs4("time")%> 
            </font></td>
        </tr>
      </table>
      <font color="#000000"> 
      <% rs4.close %>
      <% idd=rs("id") %>
      <%
	set co5=server.createobject("adodb.recordset")
	co5.open "select count(id) as sum from wenzhang where boardid="&idd&"",conn,1,3
	if not co5.eof then
%>
      </font> </td>
  </tr>
  <tr bgcolor="#FFFFFF" bordercolor="#F5F5F5"> 
    <td colspan="2" bgcolor="#F5F5F5"><font color="#000000"></font><font color="#000000"></font></td>
    <td height="10" width="451" bgcolor="#F5F5F5"><font size="2" color="#000000"><img src="img/banzhu.gif" width="18" height="17">版主：<%=rs("banzhu")%> 
      <img src="img/BMsg.gif" width="16" height="16"></font></td>
    <td height="10" width="242" bgcolor="#F5F5F5"><font color="#000000"><img src="img/elist.gif" width="12" height="13">本版帖子总数：<%=co5("sum")%> 
      <%
	end if
	co5.close
	set co5=nothing
	%>
      </font></td>
  </tr>
</table>
<% 
rs.movenext  
loop  
rs.Close  
%>
<% 
rst.movenext
loop
rst.close 
%>
<br>
<% 
dim rs2
dim sql2
set rs2=server.createobject("adodb.recordset")
sql2="select*from sum"
on error resume next
rs2.open sql2,conn,1,3
rs2("num")=rs2("num")+1
rs2.update
%>
<table width="764" border="0" style="border: 1pt solid #333333" cellspacing="0" cellpadding="0" align="center" bordercolor="#F5F5F5">
  <tr> 
    <td height="30" colspan="2" background="images/tbg.gif"><font color="#000000"><img src="img/nofollow.gif" width="15" height="15"><b><font color="#FFFFFF">童真岁月论坛&gt;&gt;&gt;在线统计&gt;&gt;&gt;</font></b> 
      <font color="#FFFFFF"> 
      <%
	set co=server.createobject("adodb.recordset")
	co.open "select count(id) as sum from wenzhang",conn,1,3
	if not co.eof then
%>
      <%
	set co1=server.createobject("adodb.recordset")
	co1.open "select count(rid) as sum from rwenzhang",conn,1,3
	if not co1.eof then
%>
      </font></font><font color="#FFFFFF"> </font></td>
  </tr>
  <tr> 
    <td height="66" width="32" bordercolor="#F5F5F5"> 
      <div align="center"><img src="img/totel.gif" width="23" height="23"></div>
    </td>
    <td height="66" width="724" bordercolor="#F5F5F5"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="60%">■目前本论坛在线共<%=server.execute("zaixian.asp")%>人，总访问量<%=rs2("num")%>人次，<a href="user.asp" 
      onClick="return pop(this.href);">在线用户详细列表</a>　　 </td>
          <td width="40%">■您的IP地址：<%=request.servervariables("remote_addr")%> 
          </td>
        </tr>
        <tr>
          <td width="60%"><font color="#000000"><font size="2">■目前论坛共有</font></font><font size="2" color="#000000">主题</font><font color="#000000"><%=co("sum")%></font><font size="2" color="#000000">篇，回复<%=co1("sum")%>篇 
            <%
	end if
	co1.close
	set co=nothing
	%>
            </font><font color="#000000"> 
            <%
	end if
	co.close
	set co=nothing
	%>
            <%
function system(info)
if Instr(info,"NT 5.1")>0 then
system=system+"Windows XP"
elseif Instr(info,"Tel")>0 then
system=system+"Telport"
elseif Instr(info,"webzip")>0 then
system=system+"webzip"
elseif Instr(info,"flashget")>0 then
system=system+"flashget"
elseif Instr(info,"offline")>0 then
system=system+"offline"
elseif Instr(info,"NT 5")>0 then
system=system+"Windows 2000"
elseif Instr(info,"NT 4")>0 then
system=system+"Windows NT4"
elseif Instr(info,"98")>0 then
system=system+"Windows 98"
elseif Instr(info,"95")>0 then
system=system+"Windows 95"
elseif instr(info,"unix") or instr(info,"linux") or instr(info,"SunOS") or instr(info,"BSD") then
system=system+"?Unix"
elseif instr(thesoft,"Mac") then
system=system+"Mac"
else
system=system+"??"
end if
end function

function browser(info)
if Instr(info,"NetCaptor 6.5.0")>0 then
browser="NetCaptor 6.5.0"
elseif Instr(info,"MyIe 3.1")>0 then
browser="MyIe 3.1"
elseif Instr(info,"NetCaptor 6.5.0RC1")>0 then
browser="NetCaptor 6.5.0RC1"
elseif Instr(info,"NetCaptor 6.5.PB1")>0 then
browser="NetCaptor 6.5.PB1"
elseif Instr(info,"MSIE 5.5")>0 then
browser="Internet Explorer 5.5"
elseif Instr(info,"MSIE 6.0")>0 then
browser="Internet Explorer 6.0"
elseif Instr(info,"MSIE 6.0b")>0 then
browser="Internet Explorer 6.0b"
elseif Instr(info,"MSIE 5.01")>0 then
browser="Internet Explorer 5.01"
elseif Instr(info,"MSIE 5.0")>0 then
browser="Internet Explorer 5.00"
elseif Instr(info,"MSIE 4.0")>0 then
browser="Internet Explorer 4.01"
else
browser="??"
end if
end function
%>
            </font><font color="#FFFFFF"><font color="#000000"> 
            <%
dim rsnew 
dim sqlnew 
set rsnew = server.createobject("adodb.recordset") 
sqlnew = "select count(id) as sum from wenzhang where time<now() and time>now()-1"
on error resume next 
rsnew.Open sqlnew,conn,1 
RSnew.MoveFirst 
do while not rsnew.eof 
%>
            今日新帖：<%=rsnew("sum")%>篇</font> 
            <% rsnew.movenext
loop
rsnew.close
%>
            </font><font color="#000000"> </font></td>
          <td width="40%"><font color="#000000"> ■您的浏览器类型:<%=browser(Request.ServerVariables("HTTP_USER_AGENT"))%></font><font color="#FFFFFF"></font></td>
        </tr>
        <tr>
          <td width="60%"><font color="#FFFFFF"><font color="#000000">■<% dim rsn 
	   dim sqln
	   set rsn=server.createobject("adodb.recordset")
	   sqln="select*from peek"
	   rsn.open sqln,conn,1
	   %>
            历史最高在线<font color="#000000"><%=rsn("peeknumber")%></font> ，发生于 <font color="#000000"> 
            </font> <font color="#000000"><%=rsn("itime")%></font> <%=Application("num")%> 
            <% peeknumber=application("num") %>
            <% if peeknumber>rsn("peeknumber") then
   conn.execute("update peek set peeknumber='"&peeknumber&"',itime=now()")%>
            <% end if %>
            <%
	rsn.close
	%>
            </font></font></td>
          <td width="40%"><font color="#000000">■您的操作系统:<%=system(Request.ServerVariables("HTTP_USER_AGENT"))%></font></td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
<% rs2.close %>
<%
strsql="select * from user where online=1 "
set rs=server.createobject("ADODB.recordset")
rs.open strsql,conn,2,1
%>
<font color="#FFFFFF">
</font><br>
<table width="762" border="0" style="border: 1pt solid #333333" cellspacing="0" cellpadding="0" align="center" bordercolor="#F5F5F5">
  <tr > 
    <td height="31" bgcolor="#F5F5F5" bordercolor="#F5F5F5" background="images/tbg.gif"><font color="#FFFFFF"><img src="img/nofollow.gif" width="15" height="15"><b>童真岁月论坛&gt;&gt;&gt;友情链接&gt;&gt;&gt; 
      </b></font><font color="#000000"> </font></td>
  </tr>
  <tr> 
    <td height="82" bgcolor="#FFFFFF" bordercolor="#F5F5F5"> 
      <% set con=server.createobject("adodb.recordset")
	con.open "select count(id) as sum from link",conn,1,3
	if not con.eof then %>
      <%
strsql="select * from link "
set rsa=server.createobject("ADODB.recordset")
rsa.open strsql,conn,2,1
%>
      <br>
      <font size="2"> </font> 
      <table width="530" align="center" border="0" height="32" cellspacing="0" cellpadding="0">
        <%for i=1 to con("sum")
if (i mod 6=1) then
response.write"<tr>"
end if
response.write"<td><a href="&rsa("siteurl")&" target=_blank><img src="&rsa("logourl")&" width=88 height=31 border=0></a></td>"
if (i mod 6=0) then
response.write"<tr></tr>"
end if
rsa.movenext
next%>
      </table>
      <font size="2">
      <%
	end if
	con.close
	set con=nothing
	%>
      </font> </td>
  </tr>
</table>
<table width="764" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="24" bgcolor="#F5F5F5"> 
      <div align="center"><font color="#000000">| 网站设计 | 友情链接 | 给我留言 | QQ：349643649 
        |</font></div>
    </td>
  </tr>
  <tr>
    <td height="18">
      <div align="center">&copy; 2006 Childhood Of Saint Bbs 童真岁月ASP官方论坛 版权所有</div>
    </td>
  </tr>
  <tr> 
    <td height="18"> 
      <div align="center"><font size="2">Childhood Of Saint Bbs 程序耗时： 
        <% t2=now() 
shijian=cstr(cdbl((t2-t1)*24*60*60))
response.write""&shijian&""&"秒" 
%>
        </font></div>
    </td>
  </tr>
</table>
</body>
</html>

