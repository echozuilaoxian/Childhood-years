<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<%
dim rs4 
dim sql4 
set rs4 = server.createobject("adodb.recordset") 
sql4 = "select*from wenzhang where id="&Trim(Request.QueryString("id"))  
rs4.Open sql4,conn,1
if rs4("lock")=1 then
response.write"<script>alert('�Բ��𣬱������Ѿ���������ֹ�ظ���');history.back(-1);</script>"
response.End()
end if
rs4.close %><html>
<head>
<title>�ظ�����--��̳ϵͳ</title>
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
		alert("Email ���\n���� Email ��������\n�÷�1: [email]nobody@domain.com[/email]\n�÷�2: [email=nobody@domain.com]����[/email]");
	} else if (basic) {
		AddTxt="[email][/email]";
		AddText(AddTxt);
	} else {
		txt2=prompt("������ʾ������.\n���Ϊ�գ���ô��ֻ��ʾ��� Email ��ַ","");
		if (txt2!=null) {
			txt=prompt("Email ��ַ.","name@domain.com");
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
		alert("Flash ����\n���� Flash ����.\n�÷�: [flash]Flash �ļ��ĵ�ַ[/flash]");
	} else if (basic) {
		AddTxt="[flash][/flash]";
		AddText(AddTxt);
	} else {
		txt=prompt("Flash �ļ��ĵ�ַ","http://");
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
		alert("Shockwave ����\n���� Shockwave ����.\n�÷�: [dir=500,350]Shockwave �ļ��ĵ�ַ[/dir]");
	} else if (basic) {
		AddTxt="[dir][/dir]";
		AddText(AddTxt);
	} else {
		txt=prompt("Shockwave �ļ��ĵ�ַ","http://");
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
		alert("real player �ļ�\n���� real player �ļ�.\n�÷�: [rm=500,350]real player �ļ��ĵ�ַ[/rm]");
	} else if (basic) {
		AddTxt="[rm][/rm]";
		AddText(AddTxt);
	} else {
		txt=prompt("real player �ļ��ĵ�ַ","http://");
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
		alert("media player �ļ�\n���� wmv �ļ�.\n�÷�: [mp=500,350]media player �ļ��ĵ�ַ[/mp]");
	} else if (basic) {
		AddTxt="[mp][/mp]";
		AddText(AddTxt);
	} else {
		txt=prompt("media player �ļ��ĵ�ַ","http://");
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
		alert("quick time �ļ�\n���� quick time �ļ�.\n�÷�: [qt=500,350]quick time �ļ��ĵ�ַ[/qt]");
	} else if (basic) {
		AddTxt="[qt][/qt]";
		AddText(AddTxt);
	} else {
		txt=prompt("quick time �ļ��ĵ�ַ","http://");
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
		alert("���ִ�С���\n�������ִ�С.\n�ɱ䷶Χ 1 - 6.\n 1 Ϊ��С 6 Ϊ���.\n�÷�: [size="+size+"]���� "+size+" ����[/size]");
	} else if (basic) {
		AddTxt="[size="+size+"][/size]";
		AddText(AddTxt);
	} else {
		txt=prompt("��С "+size,"����");
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
		alert("�Ӵֱ��\nʹ�ı��Ӵ�.\n�÷�: [b]���ǼӴֵ�����[/b]");
	} else if (basic) {
		AddTxt="[b][/b]";
		AddText(AddTxt);
	} else {
		txt=prompt("���ֽ������.","����");
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
		alert("б����\nʹ�ı������Ϊб��.\n�÷�: [i]����б����[/i]");
	} else if (basic) {
		AddTxt="[i][/i]";
		AddText(AddTxt);
	} else {
		txt=prompt("���ֽ���б��","����");
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
		alert("���ñ��\n����һЩ����.\n�÷�: [quote]��������[/quote]");
	} else if (basic) {
		AddTxt="[quote][/quote]";
		AddText(AddTxt);
	} else {
		txt=prompt("�����õ�����","����");
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
		alert("��ɫ���\n�����ı���ɫ.  �κ���ɫ�������Ա�ʹ��.\n�÷�: [color="+color+"]��ɫҪ�ı�Ϊ"+color+"������[/color]");
	} else if (basic) {
		AddTxt="[color="+color+"][/color]";
		AddText(AddTxt);
	} else {
     	txt=prompt("ѡ�����ɫ��: "+color,"����");
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
		alert("������\nʹ��������, ����ʹ�ı�����롢���С��Ҷ���.\n�÷�: [align=center|left|right]Ҫ������ı�[/align]");
	} else if (basic) {
		AddTxt="[align=center|left|right][/align]";
		AddText(AddTxt);
	} else {
		txt2=prompt("������ʽ\n���� 'center' ��ʾ����, 'left' ��ʾ�����, 'right' ��ʾ�Ҷ���.","center"); 
		while ((txt2!="") && (txt2!="center") && (txt2!="left") && (txt2!="right") && (txt2!=null)) {
			txt2=prompt("����!\n����ֻ������ 'center' �� 'left' ���� 'right'.","");
		}
		txt=prompt("Ҫ������ı�","�ı�");
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
		alert("�������ӱ��\n����һ���������ӱ��\nʹ�÷���: [url]http://www.xhonline.cn[/url]\nUSE: [url=http://www.xhonline.cn]��������[/url]");
	} else if (basic) {
		AddTxt="[url][/url]";
		AddText(AddTxt);
	} else {
		txt2=prompt("�����ı���ʾ.\n�������ʹ��, ����Ϊ��, ��ֻ��ʾ�������ӵ�ַ. ","");
		if (txt2!=null) {
			txt=prompt("��������.","http://");
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
		alert("ͼƬ���\n����ͼƬ\n�÷��� [img]http://www.xhonline.cn/logo.gif[/img]");
	} else if (basic) {
		AddTxt="[img][/img]";
		AddText(AddTxt);
	} else {
		txt=prompt("ͼƬ�� URL","http://");
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
		alert("������\nʹ�ô����ǣ�����ʹ��ĳ����������� html �ȱ�־���ᱻ�ƻ�.\nʹ�÷���:\n [code]�����Ǵ�������[/code]");
	} else if (basic) {
		AddTxt="\r[code]\r[/code]";
		AddText(AddTxt);
	} else {
		txt=prompt("�������","");
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
		alert("�б���\n����һ�����ֻ��������б�.\n\nUSE: [list] [*]��Ŀһ[/*] [*]��Ŀ��[/*] [*]��Ŀ��[/*] [/list]");
	} else if (basic) {
		AddTxt=" [list][*]  [/*][*]  [/*][*]  [/*][/list]";
		AddText(AddTxt);
	} else {
		txt=prompt("�б�����\n���� 'A' ��ʾ�����б�, '1' ��ʾ�����б�, ���ձ�ʾ�����б�.","");
		while ((txt!="") && (txt!="A") && (txt!="a") && (txt!="1") && (txt!=null)) {
			txt=prompt("����!\n����ֻ������ 'A' �� '1' ��������.","");
		}
		if (txt!=null) {
			if (txt=="") {
				AddTxt="[list]";
			} else {
				AddTxt="[list="+txt+"]";
			}
			txt="1";
			while ((txt!="") && (txt!=null)) {
				txt=prompt("�б���\n�հױ�ʾ�����б�","");
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
		alert("������\n��������������.\n�÷�: [face="+font+"]�ı���������Ϊ"+font+"[/face]");
	} else if (basic) {
		AddTxt="[face="+font+"][/face]";
		AddText(AddTxt);
	} else {
		txt=prompt("Ҫ�������������"+font,"����");
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
		alert("�»��߱��\n�����ּ��»���.\n�÷�: [u]Ҫ���»��ߵ�����[/u]");
	} else if (basic) {
		AddTxt="[u][/u]";
		AddText(AddTxt);
	} else {
		txt=prompt("�»�������.","����");
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
		alert("������\nʹ���ַ���.\n�÷�: [fly]����Ϊ��������[/fly]");
	} else if (basic) {
		AddTxt="[fly][/fly]";
		AddText(AddTxt);
	} else {
		txt=prompt("��������","����");
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
		alert("�ƶ����\nʹ���ֲ����ƶ�Ч��.\n�÷�: [move]Ҫ�����ƶ�Ч��������[/move]");
	} else if (basic) {
		AddTxt="[move][/move]";
		AddText(AddTxt);
	} else {
		txt=prompt("Ҫ�����ƶ�Ч��������","����");
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
               alert("��Ӱ���\nʹ���ֲ�����ӰЧ��.\n�÷�: [SHADOW=���, ��ɫ, �߽�]Ҫ������ӰЧ��������[/SHADOW]");
	} else if (basic) {
		AddTxt="[SHADOW=255,blue,1][/SHADOW]";
		AddText(AddTxt);
	} else {
		txt2=prompt("���ֵĳ��ȡ���ɫ�ͱ߽��С","255,blue,1");
		if (txt2!=null) {
			txt=prompt("Ҫ������ӰЧ��������","����");
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
		alert("���α��\nʹ���ֲ�������Ч��.\n�÷�: [GLOW=���, ��ɫ, �߽�]Ҫ��������Ч��������[/GLOW]");
	} else if (basic) {
		AddTxt="[glow=255,red,2][/glow]";
		AddText(AddTxt);
	} else {
		txt2=prompt("���ֵĳ��ȡ���ɫ�ͱ߽��С","255,red,2");
		if (txt2!=null) {
			txt=prompt("Ҫ��������Ч��������.","����");
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
      <a href="default.asp">��ҳ</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="reg.asp">ע��</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>��½</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">�˳�</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">����Ա���</a> 
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
    <td height="24" bgcolor="#CCCCCC"><b>ͯ��������̳--����ظ���</b></td>
  </tr>
</table>
<form name="myform" method="post" action="saver.asp" onSubmit="return LoginAdd()">
  <font size="2"> </font>
  <table width="757" border="0" align="center" bordercolor="#CCCCCC" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="141" bgcolor="#F5F5F5" height="27"> 
        <div align="right">�û����ƣ�</div>
      </td>
      <td width="606" bgcolor="#F5F5F5" height="27"> <font color="#FF0000">* 
        <input type="text" name="name" value="<%=session("user")%>" size="23" style="background-color:transparent;border:1px solid #000000">
        </font></td>
    </tr>
    <tr> 
      <td width="141" bgcolor="#F5F5F5" height="30"> 
        <div align="right">�û����</div>
      </td>
      <td width="606" bgcolor="#F5F5F5" height="30"> <font color="#FF0000">* 
        <input type="password" name="pwd" value="<%=session("pwd")%>" size="25" style="background-color:transparent;border:1px solid #000000">
        </font></td>
    </tr>
    <tr> 
      <td width="141" bgcolor="#F5F5F5" height="26"> 
        <div align="right">�ظ����⣺</div>
      </td>
      <td width="606" bgcolor="#F5F5F5" height="26"> <font color="#FF0000"> ��Re:<%=rs("title")%></font></td>
    </tr>
    <tr> 
      <td width="141" bgcolor="#F5F5F5" height="27"> 
        <div align="right">UBB���룺</div>
      </td>
      <td width="606" bgcolor="#F5F5F5" height="27"><font color="#FF0000"> * <img onClick="bold()" align="absmiddle" src="img/Ubb_bold.gif" width="22" height="22" alt="����" border="0"><img onClick="underline()" align="absmiddle" src="img/Ubb_underline.gif" width="23" height="22" alt="�»���" border="0"><img onClick="center()" align="absmiddle" src="img/Ubb_center.gif" width="22" height="22" alt="����" border="0"><img onClick="emails()" align="absmiddle" src="img/Ubb_email.gif" width="23" height="22" alt="Email����" border="0"><img onClick="hyperlink()" align="absmiddle" src="img/Ubb_url.gif" width="22" height="22" alt="��������" border="0"><img onClick="image()" align="absmiddle" src="img/Ubb_image.gif" width="23" height="22" alt="ͼƬ" border="0"><img onClick="flash()" align="absmiddle" src="img/Ubb_swf.gif" width="23" height="22" alt="FlashͼƬ" border="0"><img onClick="Cdir()" align="absmiddle" src="img/Ubb_Shockwave.gif" width="23" height="22" alt="Shockwave�ļ�" border="0"><img onClick="Cmov()" align="absmiddle" src="img/Ubb_qt.gif" width="23" height="22" alt="QuickTime��Ƶ�ļ�" border="0"><img onClick="showcode()" align="absmiddle" src="img/Ubb_code.gif" width="22" height="22" alt="����" border="0"><img onClick="quote()" align="absmiddle" src="img/Ubb_quote.gif" width="23" height="22" alt="����" border="0"><img onClick="Crm()" align="absmiddle" src="img/Ubb_rm.gif" width="23" height="22" alt="realplay��Ƶ�ļ�" border="0"><img onClick="Cwmv()" align="absmiddle" src="img/Ubb_mp.gif" width="23" height="22" alt="Media Player��Ƶ�ļ�" border="0"><img onClick="setfly()" align="absmiddle" height="22" alt="������" src="img/Ubb_fly.gif" width="23" border="0"><img onClick="move()" align="absmiddle" height="22" alt="�ƶ���" src="img/Ubb_move.gif" width="23" border="0"><img onClick="glow()" align="absmiddle" height="22" alt="������" src="img/Ubb_glow.gif" width="23" border="0"><img onClick="shadow()" align="absmiddle" height="22" alt="��Ӱ��" src="img/Ubb_shadow.gif" width="23" border="0"> 
        </font></td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#F5F5F5"> 
        <div align="right">�������ݣ�</div>
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
      <td width="606" bgcolor="#F5F5F5"> ���� 
        <input type="submit" name="Submit2" value="����ظ�">
        <input type="reset" name="Submit2" value="ȡ��">
      </td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#F5F5F5"> ����������<br>
        1�����⾡���ܼ�����Ҫ�����ݾ�����˵������飬�����ظ��߲�����ȷ����ʱ�Ļظ���<br>
        2����������֧��UBB���빦�ܣ�������ʹ���ձ��UBB��������������¸��Ӿ��ʡ�<br>
        3��������ѡ��ͬ�������־���������¡� </td>
    </tr>
  </table>
</form>
<br>
<table width="758" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td bgcolor="#F5F5F5" height="24"> 
      <div align="center"><font color="#000000">| ��վ��� | �������� | �������� | QQ��349643649 
        |</font></div>
    </td>
  </tr>
  <tr> 
    <td height="18"> 
      <div align="center">&copy; 2006 Childhood Of Saint Bbs ͯ������ASP�ٷ���̳ ��Ȩ����</div>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<% rs.close %></body>
</html>

