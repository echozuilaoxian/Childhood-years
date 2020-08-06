<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<!--#include file="ubbcode.asp" -->
<% id=request.querystring("id") %>
<% conn.execute("update wenzhang set hits=hits+1 where id="&id&"") %>
<%
dim rs5 
dim sql5 
set rs5 = server.createobject("adodb.recordset") 
sql5 = "select*from wenzhang where id="&Trim(Request.QueryString("id")) 
on error resume next 
rs5.Open sql5,conn,1 
RS5.MoveFirst 
do while not rs5.eof 
%>
<% banhao=rs5("boardid") %><html>
<%
dim rs4 
dim sql4 
set rs4 = server.createobject("adodb.recordset") 
sql4 = "select top 5 *from board where id="&banhao&""
on error resume next 
rs4.Open sql4,conn,1 
RS4.MoveFirst 
do while not rs4.eof 
%>
<head>
<title>Welcome to <%=rs4("boardname")%>-- 童真岁月论坛系统</title>
<meta http-equiv="GuestContent-Type" GuestContent="text/html; charset=gb2312"> 
<SCRIPT type=text/javascript>  
function bbimg(o)
{
  var zoom=parseInt(o.style.zoom, 10)||100;
  zoom+=event.wheelDelta/12;
  if (zoom>0) o.style.zoom=zoom+'%';
  return false;
}
</script>
<script language="javascript">
helpstat = false;
stprompt = true;
basic = false;
function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 0) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 2) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}

function AddText(NewCode) {
document.myform.GuestGuestContent.value+=NewCode;
}

function emails() {
	if (helpstat) {
		alert("Email 标记\n插入 Email 超级链接\n用法1: [email]nobody@domain.com[/email]\n用法2: [email=nobody@domain.com]佚名[/email]");
	} else if (basic) {
		AddTxt="[email][/email]";
		AddText(AddTxt);
	} else {
		txt2=prompt("链接显示的文字.\n如果为空，那么将只显示你的 Email 地址","");
		if (txt2!=null) {
			txt=prompt("Email 地址.","name@domain.com");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[email]"+txt+"[/email]";
				} else {
					AddTxt="[email="+txt+"]"+txt2;
					AddText(AddTxt);
					AddTxt="[/email]";
				}
				AddText(AddTxt);
			}
		}
	}
}

function flash() {
 	if (helpstat){
		alert("Flash 动画\n插入 Flash 动画.\n用法: [flash]Flash 文件的地址[/flash]");
	} else if (basic) {
		AddTxt="[flash][/flash]";
		AddText(AddTxt);
	} else {
		txt=prompt("Flash 文件的地址","http://");
		if (txt!=null) {
			AddTxt="[flash]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
		}
	}
}

function Cdir() {
 	if (helpstat){
		alert("Shockwave 动画\n插入 Shockwave 动画.\n用法: [dir=500,350]Shockwave 文件的地址[/dir]");
	} else if (basic) {
		AddTxt="[dir][/dir]";
		AddText(AddTxt);
	} else {
		txt=prompt("Shockwave 文件的地址","http://");
		if (txt!=null) {
			AddTxt="[dir=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/dir]";
			AddText(AddTxt);
		}
	}
}

function Crm() {
 	if (helpstat){
		alert("real player 文件\n插入 real player 文件.\n用法: [rm=500,350]real player 文件的地址[/rm]");
	} else if (basic) {
		AddTxt="[rm][/rm]";
		AddText(AddTxt);
	} else {
		txt=prompt("real player 文件的地址","http://");
		if (txt!=null) {
			AddTxt="[rm=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/rm]";
			AddText(AddTxt);
		}
	}
}

function Cwmv() {
 	if (helpstat){
		alert("media player 文件\n插入 wmv 文件.\n用法: [mp=500,350]media player 文件的地址[/mp]");
	} else if (basic) {
		AddTxt="[mp][/mp]";
		AddText(AddTxt);
	} else {
		txt=prompt("media player 文件的地址","http://");
		if (txt!=null) {
			AddTxt="[mp=1500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/mp]";
			AddText(AddTxt);
		}
	}
}

function Cmov() {
 	if (helpstat){
		alert("quick time 文件\n插入 quick time 文件.\n用法: [qt=500,350]quick time 文件的地址[/qt]");
	} else if (basic) {
		AddTxt="[qt][/qt]";
		AddText(AddTxt);
	} else {
		txt=prompt("quick time 文件的地址","http://");
		if (txt!=null) {
			AddTxt="[qt=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/qt]";
			AddText(AddTxt);
		}
	}
}


function showsize(size) {
	if (helpstat) {
		alert("文字大小标记\n设置文字大小.\n可变范围 1 - 6.\n 1 为最小 6 为最大.\n用法: [size="+size+"]这是 "+size+" 文字[/size]");
	} else if (basic) {
		AddTxt="[size="+size+"][/size]";
		AddText(AddTxt);
	} else {
		txt=prompt("大小 "+size,"文字");
		if (txt!=null) {
			AddTxt="[size="+size+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/size]";
			AddText(AddTxt);
		}
	}
}

function bold() {
	if (helpstat) {
		alert("加粗标记\n使文本加粗.\n用法: [b]这是加粗的文字[/b]");
	} else if (basic) {
		AddTxt="[b][/b]";
		AddText(AddTxt);
	} else {
		txt=prompt("文字将被变粗.","文字");
		if (txt!=null) {
			AddTxt="[b]"+txt;
			AddText(AddTxt);
			AddTxt="[/b]";
			AddText(AddTxt);
		}
	}
}

function italicize() {
	if (helpstat) {
		alert("斜体标记\n使文本字体变为斜体.\n用法: [i]这是斜体字[/i]");
	} else if (basic) {
		AddTxt="[i][/i]";
		AddText(AddTxt);
	} else {
		txt=prompt("文字将变斜体","文字");
		if (txt!=null) {
			AddTxt="[i]"+txt;
			AddText(AddTxt);
			AddTxt="[/i]";
			AddText(AddTxt);
		}
	}
}

function quote() {
	if (helpstat){
		alert("引用标记\n引用一些文字.\n用法: [quote]引用内容[/quote]");
	} else if (basic) {
		AddTxt="[quote][/quote]";
		AddText(AddTxt);
	} else {
		txt=prompt("被引用的文字","文字");
		if(txt!=null) {
			AddTxt="[quote]"+txt;
			AddText(AddTxt);
			AddTxt="[/quote]";
			AddText(AddTxt);
		}
	}
}

function showcolor(color) {
	if (helpstat) {
		alert("颜色标记\n设置文本颜色.  任何颜色名都可以被使用.\n用法: [color="+color+"]颜色要改变为"+color+"的文字[/color]");
	} else if (basic) {
		AddTxt="[color="+color+"][/color]";
		AddText(AddTxt);
	} else {
     	txt=prompt("选择的颜色是: "+color,"文字");
		if(txt!=null) {
			AddTxt="[color="+color+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/color]";
			AddText(AddTxt);
		}
	}
}

function center() {
 	if (helpstat) {
		alert("对齐标记\n使用这个标记, 可以使文本左对齐、居中、右对齐.\n用法: [align=center|left|right]要对齐的文本[/align]");
	} else if (basic) {
		AddTxt="[align=center|left|right][/align]";
		AddText(AddTxt);
	} else {
		txt2=prompt("对齐样式\n输入 'center' 表示居中, 'left' 表示左对齐, 'right' 表示右对齐.","center"); 
		while ((txt2!="") && (txt2!="center") && (txt2!="left") && (txt2!="right") && (txt2!=null)) {
			txt2=prompt("错误!\n类型只能输入 'center' 、 'left' 或者 'right'.","");
		}
		txt=prompt("要对齐的文本","文本");
		if (txt!=null) {
			AddTxt="\r[align="+txt2+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/align]";
			AddText(AddTxt);
		}
	}
}

function hyperlink() {
	if (helpstat) {
		alert("超级链接标记\n插入一个超级链接标记\n使用方法: [url]http://www.xhonline.cn[/url]\nUSE: [url=http://www.xhonline.cn]链接文字[/url]");
	} else if (basic) {
		AddTxt="[url][/url]";
		AddText(AddTxt);
	} else {
		txt2=prompt("链接文本显示.\n如果不想使用, 可以为空, 将只显示超级链接地址. ","");
		if (txt2!=null) {
			txt=prompt("超级链接.","http://");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url]"+txt;
					AddText(AddTxt);
					AddTxt="[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url="+txt+"]"+txt2;
					AddText(AddTxt);
					AddTxt="[/url]";
					AddText(AddTxt);
				}
			}
		}
	}
}

function image() {
	if (helpstat){
		alert("图片标记\n插入图片\n用法： [img]http://www.xhonline.cn/logo.gif[/img]");
	} else if (basic) {
		AddTxt="[img][/img]";
		AddText(AddTxt);
	} else {
		txt=prompt("图片的 URL","http://");
		if(txt!=null) {
			AddTxt="[img]"+txt;
			AddText(AddTxt);
			AddTxt="[/img]";
			AddText(AddTxt);
		}
	}
}

function showcode() {
	if (helpstat) {
		alert("代码标记\n使用代码标记，可以使你的程序代码里面的 html 等标志不会被破坏.\n使用方法:\n [code]这里是代码文字[/code]");
	} else if (basic) {
		AddTxt="\r[code]\r[/code]";
		AddText(AddTxt);
	} else {
		txt=prompt("输入代码","");
		if (txt!=null) {
			AddTxt="[code]"+txt;
			AddText(AddTxt);
			AddTxt="[/code]";
			AddText(AddTxt);
		}
	}
}

function list() {
	if (helpstat) {
		alert("列表标记\n建造一个文字或则数字列表.\n\nUSE: [list] [*]项目一[/*] [*]项目二[/*] [*]项目三[/*] [/list]");
	} else if (basic) {
		AddTxt=" [list][*]  [/*][*]  [/*][*]  [/*][/list]";
		AddText(AddTxt);
	} else {
		txt=prompt("列表类型\n输入 'A' 表示有序列表, '1' 表示无序列表, 留空表示无序列表.","");
		while ((txt!="") && (txt!="A") && (txt!="a") && (txt!="1") && (txt!=null)) {
			txt=prompt("错误!\n类型只能输入 'A' 、 '1' 或者留空.","");
		}
		if (txt!=null) {
			if (txt=="") {
				AddTxt="[list]";
			} else {
				AddTxt="[list="+txt+"]";
			}
			txt="1";
			while ((txt!="") && (txt!=null)) {
				txt=prompt("列表项\n空白表示结束列表","");
				if (txt!="") {
					AddTxt+="[*]"+txt+"[/*]";
				}
			}
			AddTxt+="[/list] ";
			AddText(AddTxt);
		}
	}
}

function showfont(font) {
 	if (helpstat){
		alert("字体标记\n给文字设置字体.\n用法: [face="+font+"]改变文字字体为"+font+"[/face]");
	} else if (basic) {
		AddTxt="[face="+font+"][/face]";
		AddText(AddTxt);
	} else {
		txt=prompt("要设置字体的文字"+font,"文字");
		if (txt!=null) {
			AddTxt="[face="+font+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/face]";
			AddText(AddTxt);
		}
	}
}
function underline() {
  	if (helpstat) {
		alert("下划线标记\n给文字加下划线.\n用法: [u]要加下划线的文字[/u]");
	} else if (basic) {
		AddTxt="[u][/u]";
		AddText(AddTxt);
	} else {
		txt=prompt("下划线文字.","文字");
		if (txt!=null) {
			AddTxt="[u]"+txt;
			AddText(AddTxt);
			AddTxt="[/u]";
			AddText(AddTxt);
		}
	}
}
function setfly() {
 	if (helpstat){
		alert("飞翔标记\n使文字飞行.\n用法: [fly]文字为这样文字[/fly]");
	} else if (basic) {
		AddTxt="[fly][/fly]";
		AddText(AddTxt);
	} else {
		txt=prompt("飞翔文字","文字");
		if (txt!=null) {
			AddTxt="[fly]"+txt;
			AddText(AddTxt);
			AddTxt="[/fly]";
			AddText(AddTxt);
		}
	}
}

function move() {
	if (helpstat) {
		alert("移动标记\n使文字产生移动效果.\n用法: [move]要产生移动效果的文字[/move]");
	} else if (basic) {
		AddTxt="[move][/move]";
		AddText(AddTxt);
	} else {
		txt=prompt("要产生移动效果的文字","文字");
		if (txt!=null) {
			AddTxt="[move]"+txt;
			AddText(AddTxt);
			AddTxt="[/move]";
			AddText(AddTxt);
		}
	}
}

function shadow() {
	if (helpstat) {
               alert("阴影标记\n使文字产生阴影效果.\n用法: [SHADOW=宽度, 颜色, 边界]要产生阴影效果的文字[/SHADOW]");
	} else if (basic) {
		AddTxt="[SHADOW=255,blue,1][/SHADOW]";
		AddText(AddTxt);
	} else {
		txt2=prompt("文字的长度、颜色和边界大小","255,blue,1");
		if (txt2!=null) {
			txt=prompt("要产生阴影效果的文字","文字");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[SHADOW=255, blue, 1]"+txt;
					AddText(AddTxt);
					AddTxt="[/SHADOW]";
					AddText(AddTxt);
				} else {
					AddTxt="[SHADOW="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/SHADOW]";
					AddText(AddTxt);
				}
			}
		}
	}
}

function glow() {
	if (helpstat) {
		alert("光晕标记\n使文字产生光晕效果.\n用法: [GLOW=宽度, 颜色, 边界]要产生光晕效果的文字[/GLOW]");
	} else if (basic) {
		AddTxt="[glow=255,red,2][/glow]";
		AddText(AddTxt);
	} else {
		txt2=prompt("文字的长度、颜色和边界大小","255,red,2");
		if (txt2!=null) {
			txt=prompt("要产生光晕效果的文字.","文字");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[glow=255,red,2]"+txt;
					AddText(AddTxt);
					AddTxt="[/glow]";
					AddText(AddTxt);
				} else {
					AddTxt="[glow="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/glow]";
					AddText(AddTxt);
				}
			}
		}
	}
}

function openemot()
{
	var Win =window.open("guestselect.asp?action=emot","face","width=380,height=300,resizable=1,scrollbars=1");
}
function openhelp()
{
	var Win =window.open("ubbhelp.asp","face","width=550,height=400,resizable=1,scrollbars=1");
}

</script>
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
    <td height="30" bgcolor="#CCCCCC" background="images/tbg.gif"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%> 
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
<br>
<table width="763" border="0" cellspacing="0" cellpadding="0" align="center" height="23">
  <tr> 
    <td height="2"> 
      <form name="form1" method="post" action="loginok.asp">
        <table width="761" border="0" height="19" style="border: 1pt solid #333333" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="28" bgcolor="#CCCCCC" width="761" background="images/tbg.gif"><font color="#FFFFFF"><b><img src="img/nofollow.gif" width="15" height="15">快速登陆入口：</b></font></td>
          </tr>
          <tr> 
            <td height="36" bgcolor="#FFFFFF" width="761"> 
              <div align="left"><font color="#000000"> <img src="img/banzhu.gif" width="19" height="17"> 
                用户名： 
                <input type="text" name="user" size="10" style="background-color:transparent;border:1px solid #000000">
                用户： 
                <input type="text" name="pwd2" size="10" style="background-color:transparent;border:1px solid #000000">
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
      <div align="right"><font color="#000000">论坛共有注册会员：</font><font color="#0099FF"><%=co5("sum")%><font color="#333333">位</font> 
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
<table width="763" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="29"><font color="#000000"><img src="img/nofollow.gif" width="15" height="15"><font size="2"><a href="default.asp"><b>童真岁月论坛</b></a><b>&gt;&gt;&gt;所有版块&gt;&gt;&gt; 
      <a href="board.asp?id=<%=rs4("id")%>"><%=rs4("boardname")%></a> &gt;&gt;&gt;</b></font><b><font color="#000000"><font size="2"><%=server.htmlencode(rs5("title"))%></font></font></b></font></td>
  </tr>
</table>
<font size="2" color="#FFFFFF"> </font> 
<% 
rs4.movenext  
loop  
rs4.Close  
%>
<table width="763" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#333333">
  <tr bgcolor="#FFFFFF"> 
    <td height="28" width="136" bgcolor="#CCCCCC" background="images/tbg.gif">&nbsp;</td>
    <td width="313" height="28" bgcolor="#CCCCCC" background="images/tbg.gif"><font size="2" color="#FFFFFF">『<%=server.htmlencode(rs5("title"))%>』</font></td>
    <td width="301" height="28" bgcolor="#CCCCCC" background="images/tbg.gif"><font size="2" color="#FFFFFF">发表于：<%=rs5("time")%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td rowspan="2" height="47" width="136" valign="top"><font size="2"></font> 
      <div align="center"><font size="2" color="#003399">『<%=rs5("name")%>』</font> 
        <% author=rs5("name") %>
        <%
dim rs6 
dim sql6 
set rs6 = server.createobject("adodb.recordset") 
sql6 = "select*from user where user='"&author&"'" 
on error resume next 
rs6.Open sql6,conn,1 
RS6.MoveFirst 
do while not rs6.eof
%>
        <br>
      </div>
      <table width="117" border="0" cellspacing="0" cellpadding="0" align="center" height="96">
        <tr> 
          <td height="42"> 
            <div align="center"><img src="images/<%=rs6("userico")%>.jpg" border="1"></div>
          </td>
        </tr>
        <tr> 
          <td><font size="2">昵称：<%=rs6("user")%></font></td>
        </tr>
        <tr> 
          <td>星级： 
            <% if rs6("GDP")<100 then
		  response.write"<font color=red>☆</font>☆☆☆☆☆"
		  end if
		  if rs6("GDP")<200 and rs6("GDP")>100 then
		  response.write"<font color=red>☆☆</font>☆☆☆☆"
		  end if
		  if rs6("GDP")<300 and rs6("GDP")>200 then
		  response.write"<font color=red>☆☆☆</font>☆☆☆"
		  end if
		  if rs6("GDP")<400 and rs6("GDP")>300 then
		  response.write"<font color=red>☆☆☆☆</font>☆☆"
		  end if 
		  if rs6("GDP")<500 and rs6("GDP")>400 then
		  response.write"<font color=red>☆☆☆☆☆</font>☆"
		  end if
		  if rs6("GDP")>500 then
		  response.write"<font color=red>☆☆☆☆☆☆</font>"
		  end if %>
          </td>
        </tr>
        <tr> 
          <td><font size="2">身份：<%=rs6("sex")%></font></td>
        </tr>
        <tr> 
          <td>等级： 
            <% if rs6("GDP")>200 then
response.write"高级大法师"
end if
if rs6("GDP")<200 and rs6("GDP")>100 then
response.write"中级法师"
end if
if rs6("GDP")<100 then
response.write"魔法学徒"
end if %>
          </td>
        </tr>
        <tr> 
          <td><font size="2">经验：<font color="#FF0000"><%=rs6("GDP")%>分</font></font></td>
        </tr>
        <tr> 
          <td><font size="2">注册：<%=rs6("time")%></font></td>
        </tr>
      </table>
      <br>
    </td>
    <td height="153" colspan="2"><font size="2"> </font> 
      <table width="619" border="0" cellpadding="0" cellspacing="0" height="152" >
        <tr> 
          <td height="15"> 
            <div align="right"><a href="info.asp?user=<%=rs6("user")%>"><img src="images/info.gif" width="45" height="18" border="0" alt="查看<%=rs6("user")%>的个人资料"></a> 
              <a href="userarticle.asp?user=<%=rs6("user")%>"><img src="images/find.gif" width="45" height="18" border="0" alt="搜索<%=rs6("user")%>的相关文章"></a> 
              <a href="articlemodify.asp?id=<%=rs5("id")%>"><img src="images/edit.gif" width="48" height="18" border="0" alt="重新编辑文章"></a> 
              <a href="reply.asp?id=<%=rs5("id")%>"><img src="images/reply.gif" width="45" height="18" alt="回复本主题或文章" border="0"></a></div>
          </td>
        </tr>
        <tr> 
          <td height="58"><font size="2" color="#000000"><b><%=server.htmlencode(rs5("title"))%></b></font> 
            <table width="613" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed;word-break:break-all">
              <tr> 
                <td><font size="2"><%=UBBCode(dvHTMLEncode(rs5("content")))%></font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td height="2"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0" height="25">
              <tr> 
                <td height="11"> 
                  <div align="center"><img src="images/sigline.gif" width="363" height="16"></div>
                </td>
              </tr>
              <tr> 
                <td> 
                  <div align="center"><font size="2" color="#000000">个性签名： ※</font><%=UBBCode(dvHTMLEncode(rs6("sign")))%><font size="2" color="#000000">※</font></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="321" height="20" bgcolor="#FFFFFF"><font size="2"><img src="img/oicq.gif" width="16" height="16">:<%=rs6("QQ")%> 
      　<img src="img/status_5.gif" width="14" height="11">:<%=rs6("email")%> 
      <%       
rs6.movenext  
loop    
rs6.Close  
%>
      </font></td>
    <td width="301" height="20" bgcolor="#FFFFFF"><font size="2"><img src="img/ip.gif" width="13" height="15">IP:<%=rs5("ip")%></font></td>
  </tr>
</table>
<% idd=Request.Querystring("id") %>
<%
dim rs
dim sql
set rs = server.createobject("adodb.recordset")
sql = "select*from rwenzhang where bid=0 and rid="&idd&""
count=conn.execute("select count(id) from rwenzhang where bid=0 and rid="&idd&"")(0)
on error resume next
pagesetup=8
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
<table width="763" align="center" border="0" cellpadding="0" cellspacing="1" bgcolor="#333333" style="word-break:break-all;border-top-width:0px">
  <tr> 
    <td width="137" height="56" rowspan="3" valign="top" bgcolor="#FFFFFF"> 
      <div align="left"> </div>
      <div align="center"><font size="2"> </font><font size="2" color="#003399">『</font><font size="2"><%=rs("rname")%></font><font size="2" color="#003399">』</font><font size="2"> 
        <% rauthor=rs("rname") %>
        </font><font size="2"> 
        <%
dim rs7 
dim sql7 
set rs7 = server.createobject("adodb.recordset") 
sql7 = "select*from user where user='"&rauthor&"'" 
on error resume next 
rs7.Open sql7,conn,1 
RS7.MoveFirst 
do while not rs7.eof
%>
        </font><font size="2" color="#003399"><br>
        </font></div>
      <table width="117" border="0" cellspacing="0" cellpadding="0" align="center" height="36">
        <tr> 
          <td height="42"> 
            <div align="center"><img src="images/<%=rs7("userico")%>.jpg" border="1"></div>
          </td>
        </tr>
        <tr> 
          <td><font size="2">昵称：<%=rs7("user")%></font></td>
        </tr>
        <tr> 
          <td>星级： 
            <% if rs7("GDP")<100 then
		  response.write"<font color=red>☆</font>☆☆☆☆☆"
		  end if
		  if rs7("GDP")<200 and rs7("GDP")>100 then
		  response.write"<font color=red>☆☆</font>☆☆☆☆"
		  end if
		  if rs7("GDP")<300 and rs7("GDP")>200 then
		  response.write"<font color=red>☆☆☆</font>☆☆☆"
		  end if
		  if rs7("GDP")<400 and rs7("GDP")>300 then
		  response.write"<font color=red>☆☆☆☆</font>☆☆"
		  end if 
		  if rs7("GDP")<500 and rs7("GDP")>400 then
		  response.write"<font color=red>☆☆☆☆☆</font>☆"
		  end if
		  if rs7("GDP")>500 then
		  response.write"<font color=red>☆☆☆☆☆☆</font>"
		  end if %>
          </td>
        </tr>
        <tr> 
          <td><font size="2">身份：<%=rs7("sex")%></font></td>
        </tr>
        <tr> 
          <td>等级： 
            <% if rs7("GDP")>200 then
response.write"高级大法师"
end if
if rs7("GDP")<200 and rs7("GDP")>100 then
response.write"中级法师"
end if
if rs7("GDP")<100 then
response.write"魔法学徒"
end if %>
          </td>
        </tr>
        <tr> 
          <td><font size="2">经验：<font color="#FF0000"><%=rs7("GDP")%>分</font></font></td>
        </tr>
        <tr> 
          <td height="3"><font size="2">注册：<%=rs7("time")%></font></td>
        </tr>
      </table>
      <font size="2"> </font><br>
      <font size="2"> </font> <font size="2"> </font></td>
    <td width="313" height="22" bgcolor="#CCCCCC" background="images/tbg.gif"> 
      <div align="left"><font size="2" color="#FFFFFF">Re:<%=rs5("title")%> </font></div>
    </td>
    <td colspan="2" height="22" width="301" bgcolor="#CCCCCC" background="images/tbg.gif"><font size="2" color="#FFFFFF">发表于：<%=rs("rtime")%></font></td>
  </tr>
  <tr> 
    <td height="95" colspan="3" bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="20"> 
            <div align="right"><a href="info.asp?user=<%=rs7("user")%>"><img src="images/info.gif" width="45" height="18" border="0" alt="查看<%=rs7("user")%>的个人资料"></a> 
              <a href="userarticle.asp?user=<%=rs7("user")%>"><img src="images/find.gif" width="45" height="18" border="0" alt="搜索<%=rs7("user")%>的相关文章"></a> 
              <a href="replymodify.asp?id=<%=rs("id")%>"><img src="images/edit.gif" width="48" height="18" border="0" alt="重新编辑文章"></a> 
            </div>
          </td>
        </tr>
        <tr> 
          <td height="114"> 
            <table width="619" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed;word-break:break-all" height="109">
              <tr> 
                <td><font size="2"><%=ubbcode(rs("rcontent"))%></font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td height="2"> 
            <table width="100%" border="0" cellpadding="0" cellspacing="0" height="25">
              <tr> 
                <td> 
                  <div align="center"><img src="images/sigline.gif" width="363" height="16"></div>
                </td>
              </tr>
              <tr> 
                <td height="7"> 
                  <div align="center"><font size="2" color="#000000">个性签名： ※</font><%=dvHTMLEncode(rs7("sign"))%><font size="2" color="#000000">※</font></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="18" width="321" bgcolor="#FFFFFF"><font size="2"><img src="img/oicq.gif" width="16" height="16">:<%=rs7("QQ")%><img src="img/status_5.gif" width="14" height="11">: 
      <%=rs7("email")%> 
      <% 
rs7.movenext  
loop    
rs7.Close  
%>
      </font></td>
    <td height="18" colspan="2" width="301" bgcolor="#FFFFFF"><font size="2"><img src="img/ip.gif" width="13" height="15">IP:<%=rs("rip")%></font></td>
  </tr>
</table>
<%i=i+1 
if i>pagesetup then exit do 
rs.movenext 
loop 
rs.Close 
%>
<table width="200" border="0" cellspacing="0" cellpadding="0" align="right">
  <tr> 
    <td> 
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
<br>
<form name="myform" method="post" action="saver.asp" onSubmit="return LoginAdd()">
  <font size="2"> 
  <% 
rs5.movenext  
loop  
rs5.Close  
%>
  </font><br>
  <table width="760" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
    <tr bgcolor="#FFFFFF"> 
      <td width="143"> 
        <div align="right">用户名称：</div>
      </td>
      <td width="612"> <font color="#FF0000">* 
        <input type="text" name="name" value="<%=session("user")%>" style="background-color:transparent;border:1px solid #000000">
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="143"> 
        <div align="right">用户口令：</div>
      </td>
      <td width="612"> <font color="#FF0000">*</font> 
        <input type="password" name="pwd" value="<%=session("pwd")%>" size="30" style="background-color:transparent;border:1px solid #000000">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="143"> 
        <div align="right">话题心情：</div>
      </td>
      <td width="612"> <font color="#FF0000">* 
        <input type="radio" name="radiobutton" value="radiobutton" checked>
        <img src="img/1.gif" width="19" height="19"> 
        <input type="radio" name="radiobutton" value="radiobutton">
        <img src="img/2.gif" width="19" height="19"> 
        <input type="radio" name="radiobutton" value="radiobutton">
        <img src="img/4.gif" width="19" height="19"> 
        <input type="radio" name="radiobutton" value="radiobutton">
        <img src="img/9.gif" width="19" height="19"> </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="143"> 
        <div align="right">UBB代码：</div>
      </td>
      <td width="612"><font color="#FF0000"> *　<img onClick="bold()" align="absmiddle" src="img/Ubb_bold.gif" width="22" height="22" alt="粗体" border="0"><img onClick="underline()" align="absmiddle" src="img/Ubb_underline.gif" width="23" height="22" alt="下划线" border="0"><img onClick="center()" align="absmiddle" src="img/Ubb_center.gif" width="22" height="22" alt="居中" border="0"><img onClick="emails()" align="absmiddle" src="img/Ubb_email.gif" width="23" height="22" alt="Email连接" border="0"><img onClick="hyperlink()" align="absmiddle" src="img/Ubb_url.gif" width="22" height="22" alt="超级连接" border="0"><img onClick="image()" align="absmiddle" src="img/Ubb_image.gif" width="23" height="22" alt="图片" border="0"><img onClick="flash()" align="absmiddle" src="img/Ubb_swf.gif" width="23" height="22" alt="Flash图片" border="0"><img onClick="Cdir()" align="absmiddle" src="img/Ubb_Shockwave.gif" width="23" height="22" alt="Shockwave文件" border="0"><img onClick="Cmov()" align="absmiddle" src="img/Ubb_qt.gif" width="23" height="22" alt="QuickTime视频文件" border="0"><img onClick="showcode()" align="absmiddle" src="img/Ubb_code.gif" width="22" height="22" alt="代码" border="0"><img onClick="quote()" align="absmiddle" src="img/Ubb_quote.gif" width="23" height="22" alt="引用" border="0"><img onClick="Crm()" align="absmiddle" src="img/Ubb_rm.gif" width="23" height="22" alt="realplay视频文件" border="0"><img onClick="Cwmv()" align="absmiddle" src="img/Ubb_mp.gif" width="23" height="22" alt="Media Player视频文件" border="0"><img onClick="setfly()" align="absmiddle" height="22" alt="飞行字" src="img/Ubb_fly.gif" width="23" border="0"><img onClick="move()" align="absmiddle" height="22" alt="移动字" src="img/Ubb_move.gif" width="23" border="0"><img onClick="glow()" align="absmiddle" height="22" alt="发光字" src="img/Ubb_glow.gif" width="23" border="0"><img onClick="shadow()" align="absmiddle" height="22" alt="阴影字" src="img/Ubb_shadow.gif" width="23" border="0"> 
        </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="143" height="25"> 
        <div align="right">话题内容：</div>
      </td>
      <td rowspan="2"> <font color="#FF0000">* 
        <textarea cols="70" rows="7" id="GuestContent" name="GuestGuestContent" style="background-color:transparent;border:1px solid #000000"></textarea>
        <input type="hidden" name="id" value="<%=request.querystring("id")%>">
        <input type="hidden" name="ip" value="<%=request.servervariables("remote_addr")%>">
        <br>
        </font> 
        <table width="572" border="0" cellspacing="0" cellpadding="0" height="13" align="center">
          <tr> 
            <td><img src="images/emot/emot01.gif" width="20" height="20" border="0" onClick="GuestGuestGuestContent.value=GuestGuestGuestContent.value+'[emot01]'"><img src="images/emot/emot02.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot02]'"><img src="images/emot/emot03.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot03]'"><img src="images/emot/emot04.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot04]'"><img src="images/emot/emot05.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot05]'"><img src="images/emot/emot06.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot06]'"><img src="images/emot/emot07.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot07]'"><img src="images/emot/emot08.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot08]'"><img src="images/emot/emot09.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot09]'"><img src="images/emot/emot10.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot10]'"><img src="images/emot/emot11.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot11]'"><img src="images/emot/emot12.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot12]'"><img src="images/emot/emot13.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot13]'"><img src="images/emot/emot14.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot14]'"><img src="images/emot/emot15.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot15]'"><img src="images/emot/emot16.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestContent.value+'[emot16]'"><img src="images/emot/emot17.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot17]'"><img src="images/emot/emot18.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot18]'"><img src="images/emot/emot19.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot19]'"><img src="images/emot/emot20.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot20]'"><img src="images/emot/emot21.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot21]'"><br>
              <img src="images/emot/emot22.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot22]'"><img src="images/emot/emot23.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot23]'"><img src="images/emot/emot24.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot24]'"><img src="images/emot/emot25.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot25]'"><img src="images/emot/emot26.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot26]'"><img src="images/emot/emot27.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot27]'"><img src="images/emot/emot28.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot28]'"><img src="images/emot/emot29.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot29]'"><img src="images/emot/emot30.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot30]'"><img src="images/emot/emot31.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot31]'"><img src="images/emot/emot32.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot32]'"><img src="images/emot/emot33.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot33]'"><img src="images/emot/emot34.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot34]'"><img src="images/emot/emot35.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot35]'"><img src="images/emot/emot36.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot36]'"><img src="images/emot/emot37.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot37]'"><img src="images/emot/emot38.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot38]'"><img src="images/emot/emot39.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot39]'"><img src="images/emot/emot40.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot40]'"><img src="images/emot/emot41.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot41]'"><img src="images/emot/emot42.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot42]'"><br>
              <img src="images/emot/emot43.gif" width="44" height="46" onClick="GuestContent.value=GuestContent.value+'[emot43]'"><img src="images/emot/emot44.gif" width="45" height="47" onClick="GuestContent.value=GuestContent.value+'[emot44]'"><img src="images/emot/emot45.gif" width="45" height="47" onClick="GuestContent.value=GuestContent.value+'[emot45]'"><img src="images/emot/emot46.gif" width="44" height="46" onClick="GuestContent.value=GuestContent.value+'[emot46]'"><img src="images/emot/emot47.gif" width="42" height="46" onClick="GuestContent.value=GuestContent.value+'[emot47]'"><img src="images/emot/emot48.gif" width="45" height="47" onClick="GuestContent.value=GuestContent.value+'[emot48]'"><img src="images/emot/emot49.gif" width="44" height="46" onClick="GuestContent.value=GuestContent.value+'[emot49]'"><img src="images/emot/emot50.gif" width="44" height="46" onClick="GuestContent.value=GuestContent.value+'[emot50]'"></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="143"> 
        <div align="right"></div>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="143">&nbsp;</td>
      <td width="612"> 
        <input type="submit" name="Submit2" value="发表回复">
        <input type="reset" name="Submit2" value="取消">
      </td>
    </tr>
  </table>
</form>
<br>
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
</body>
</html>

