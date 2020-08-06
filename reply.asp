<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<%
dim rs4 
dim sql4 
set rs4 = server.createobject("adodb.recordset") 
sql4 = "select*from wenzhang where id="&Trim(Request.QueryString("id"))  
rs4.Open sql4,conn,1
if rs4("lock")=1 then
response.write"<script>alert('对不起，本主题已经锁定，禁止回复！');history.back(-1);</script>"
response.End()
end if
rs4.close %><html>
<head>
<title>回复主题--论坛系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
document.myform.GuestContent.value+=NewCode;
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
    <td height="27" bgcolor="#CCCCCC"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%> 
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
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>登陆</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">退出</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">搜索</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">留言</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">管理员入口</a> 
      </font></td>
  </tr>
</table>
<br><% idd=request.querystring("id") %>
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select* from wenzhang where id="&idd&""
rs.open sql,conn,1
%>
<table width="757" border="0" align="center">
  <tr>
    <td height="24" bgcolor="#CCCCCC"><b>童真岁月论坛--发表回复：</b></td>
  </tr>
</table>
<form name="myform" method="post" action="saver.asp" onSubmit="return LoginAdd()">
  <font size="2"> </font>
  <table width="757" border="0" align="center" bordercolor="#CCCCCC" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="141" bgcolor="#F5F5F5" height="27"> 
        <div align="right">用户名称：</div>
      </td>
      <td width="606" bgcolor="#F5F5F5" height="27"> <font color="#FF0000">* 
        <input type="text" name="name" value="<%=session("user")%>" size="23" style="background-color:transparent;border:1px solid #000000">
        </font></td>
    </tr>
    <tr> 
      <td width="141" bgcolor="#F5F5F5" height="30"> 
        <div align="right">用户口令：</div>
      </td>
      <td width="606" bgcolor="#F5F5F5" height="30"> <font color="#FF0000">* 
        <input type="password" name="pwd" value="<%=session("pwd")%>" size="25" style="background-color:transparent;border:1px solid #000000">
        </font></td>
    </tr>
    <tr> 
      <td width="141" bgcolor="#F5F5F5" height="26"> 
        <div align="right">回复主题：</div>
      </td>
      <td width="606" bgcolor="#F5F5F5" height="26"> <font color="#FF0000"> 　Re:<%=rs("title")%></font></td>
    </tr>
    <tr> 
      <td width="141" bgcolor="#F5F5F5" height="27"> 
        <div align="right">UBB代码：</div>
      </td>
      <td width="606" bgcolor="#F5F5F5" height="27"><font color="#FF0000"> * <img onClick="bold()" align="absmiddle" src="img/Ubb_bold.gif" width="22" height="22" alt="粗体" border="0"><img onClick="underline()" align="absmiddle" src="img/Ubb_underline.gif" width="23" height="22" alt="下划线" border="0"><img onClick="center()" align="absmiddle" src="img/Ubb_center.gif" width="22" height="22" alt="居中" border="0"><img onClick="emails()" align="absmiddle" src="img/Ubb_email.gif" width="23" height="22" alt="Email连接" border="0"><img onClick="hyperlink()" align="absmiddle" src="img/Ubb_url.gif" width="22" height="22" alt="超级连接" border="0"><img onClick="image()" align="absmiddle" src="img/Ubb_image.gif" width="23" height="22" alt="图片" border="0"><img onClick="flash()" align="absmiddle" src="img/Ubb_swf.gif" width="23" height="22" alt="Flash图片" border="0"><img onClick="Cdir()" align="absmiddle" src="img/Ubb_Shockwave.gif" width="23" height="22" alt="Shockwave文件" border="0"><img onClick="Cmov()" align="absmiddle" src="img/Ubb_qt.gif" width="23" height="22" alt="QuickTime视频文件" border="0"><img onClick="showcode()" align="absmiddle" src="img/Ubb_code.gif" width="22" height="22" alt="代码" border="0"><img onClick="quote()" align="absmiddle" src="img/Ubb_quote.gif" width="23" height="22" alt="引用" border="0"><img onClick="Crm()" align="absmiddle" src="img/Ubb_rm.gif" width="23" height="22" alt="realplay视频文件" border="0"><img onClick="Cwmv()" align="absmiddle" src="img/Ubb_mp.gif" width="23" height="22" alt="Media Player视频文件" border="0"><img onClick="setfly()" align="absmiddle" height="22" alt="飞行字" src="img/Ubb_fly.gif" width="23" border="0"><img onClick="move()" align="absmiddle" height="22" alt="移动字" src="img/Ubb_move.gif" width="23" border="0"><img onClick="glow()" align="absmiddle" height="22" alt="发光字" src="img/Ubb_glow.gif" width="23" border="0"><img onClick="shadow()" align="absmiddle" height="22" alt="阴影字" src="img/Ubb_shadow.gif" width="23" border="0"> 
        </font></td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#F5F5F5"> 
        <div align="right">话题内容：</div>
        <div align="right"></div>
      </td>
      <td width="606" bgcolor="#F5F5F5"> <font color="#FF0000">* 
        <textarea cols="70" rows="7" id="content" name="GuestContent" style="background-color:transparent;border:1px solid #000000"></textarea>
        <input type="hidden" name="id" value="<%=request.querystring("id")%>">
        <input type="hidden" name="ip" value="<%=request.servervariables("remote_addr")%>">
        <br>
        </font>
        <table width="580" border="0" cellspacing="0" cellpadding="0" height="13" align="center">
          <tr> 
            <td><img src="images/emot/01.gif" width="20" height="20" border="0" onClick="GuestGuestGuestContent.value=GuestGuestGuestContent.value+'[emot01]'"><img src="images/emot/02.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot02]'"><img src="images/emot/03.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot03]'"><img src="images/emot/04.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot04]'"><img src="images/emot/05.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot05]'"><img src="images/emot/06.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot06]'"><img src="images/emot/07.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot07]'"><img src="images/emot/08.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot08]'"><img src="images/emot/09.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot09]'"><img src="images/emot/10.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot10]'"><img src="images/emot/11.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot11]'"><img src="images/emot/12.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot12]'"><img src="images/emot/13.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot13]'"><img src="images/emot/14.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot14]'"><img src="images/emot/15.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestGuestContent.value+'[emot15]'"><img src="images/emot/16.gif" width="20" height="20" onClick="GuestGuestContent.value=GuestContent.value+'[emot16]'"><img src="images/emot/17.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot17]'"><img src="images/emot/18.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot18]'"><img src="images/emot/19.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot19]'"><img src="images/emot/20.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot20]'"><img src="images/emot/21.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot21]'"><br>
              <img src="images/emot/22.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot22]'"><img src="images/emot/23.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot23]'"><img src="images/emot/24.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot24]'"><img src="images/emot/25.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot25]'"><img src="images/emot/26.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot26]'"><img src="images/emot/27.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot27]'"><img src="images/emot/28.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot28]'"><img src="images/emot/29.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot29]'"><img src="images/emot/30.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot30]'"><img src="images/emot/31.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot31]'"><img src="images/emot/32.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot32]'"><img src="images/emot/33.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot33]'"><img src="images/emot/34.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot34]'"><img src="images/emot/35.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot35]'"><img src="images/emot/36.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot36]'"><img src="images/emot/37.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot37]'"><img src="images/emot/38.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot38]'"><img src="images/emot/39.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot39]'"><img src="images/emot/40.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot40]'"><img src="images/emot/41.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot41]'"><img src="images/emot/42.gif" width="20" height="20" onClick="GuestContent.value=GuestContent.value+'[emot42]'"><br>
              <img src="images/emot/43.gif" width="44" height="46" onClick="GuestContent.value=GuestContent.value+'[emot43]'"><img src="images/emot/44.gif" width="45" height="47" onClick="GuestContent.value=GuestContent.value+'[emot44]'"><img src="images/emot/45.gif" width="45" height="47" onClick="GuestContent.value=GuestContent.value+'[emot45]'"><img src="images/emot/46.gif" width="44" height="46" onClick="GuestContent.value=GuestContent.value+'[emot46]'"><img src="images/emot/47.gif" width="42" height="46" onClick="GuestContent.value=GuestContent.value+'[emot47]'"><img src="images/emot/48.gif" width="45" height="47" onClick="GuestContent.value=GuestContent.value+'[emot48]'"><img src="images/emot/49.gif" width="44" height="46" onClick="GuestContent.value=GuestContent.value+'[emot49]'"><img src="images/emot/50.gif" width="44" height="46" onClick="GuestContent.value=GuestContent.value+'[emot50]'"></td>
          </tr>
        </table>
        <font color="#FF0000"> </font></td>
    </tr>
    <tr> 
      <td width="141" bgcolor="#F5F5F5">&nbsp;</td>
      <td width="606" bgcolor="#F5F5F5"> 　　 
        <input type="submit" name="Submit2" value="发表回复">
        <input type="reset" name="Submit2" value="取消">
      </td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#F5F5F5"> 发帖帮助：<br>
        1、标题尽可能简明扼要，内容尽可能说清楚事情，这样回复者才能明确并及时的回复。<br>
        2、文章内容支持UBB编码功能，您可以使用普遍的UBB代码帮助您的文章更加精彩。<br>
        3、您可以选择不同的心情标志来发表文章。 </td>
    </tr>
  </table>
</form>
<br>
<table width="758" border="0" cellspacing="0" cellpadding="0" align="center">
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
<p>&nbsp;</p>
<% rs.close %></body>
</html>

