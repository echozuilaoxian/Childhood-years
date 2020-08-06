<%@ CODEPAGE ="936"%>
<html>
<head>
<title>用户注册--童真岁月论坛！</title>
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
    <td height="27" bgcolor="#CCCCCC"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%><font color="#000000"> 
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
<table width="471" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399">
  <tr> 
    <td bgcolor="#F5F5F5" height="24">首页 &gt;&gt;&gt; 新用户注册 </td>
  </tr>
</table>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="467" align="center">
  <tr> 
    <td height="228"> 　 <b>继续注册前请先阅读论坛协议 </b> 
      <p> 欢迎您加入本站点参加交流和讨论，本站点为公共论坛，为维护网上公共秩序和社会稳定，请您自觉遵守以下条款：</p>
      <p>一、不得利用本站危害国家安全、泄露国家秘密，不得侵犯国家社会集体的和公民的合法权益，不得利<br>
        用本站制作、复制和传播下列信息： <br>
        　　（一）煽动抗拒、破坏宪法和法律、行政法规实施的；<br>
        　　（二）煽动颠覆国家政权，推翻社会主义制度的；<br>
        　　（三）煽动分裂国家、破坏国家统一的；<br>
        　　（四）煽动民族仇恨、民族歧视，破坏民族团结的；<br>
        　　（五）捏造或者歪曲事实，散布谣言，扰乱社会秩序的； <br>
        　　（六）宣扬封建迷信、淫秽、色情、赌博、暴力、凶杀、恐怖、教唆犯罪的；<br>
        　　（七）公然侮辱他人或者捏造事实诽谤他人的，或者进行其他恶意攻击的；<br>
        　　（八）损害国家机关信誉的；<br>
        　　（九）其他违反宪法和法律行政法规的；<br>
        　　（十）进行商业广告行为的。<br>
        二、互相尊重，对自己的言论和行为负责。<br>
      </p>
      <p>　 [<a href="regtwo.asp">同意</a>]
    </td>
  </tr>
</table>
<br>
<table width="775" border="0" cellspacing="0" cellpadding="0" align="center">
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
</table>
</body>
</html>
