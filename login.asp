<%@ CODEPAGE ="936"%>
<html>
<head>
<title>��½��̳--ͯ��������̳��</title>
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
      <a href="default.asp">��ҳ</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="reg.asp">ע��</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>��½</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">�˳�</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="http://my.phome.cn/knight999/guest/index.asp" target="_blank">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">����Ա���</a> 
      </font></td>
  </tr>
</table>
<br>
<table width="302" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399" height="19">
  <tr> 
    <td height="24"><font color="#FFFFFF"><img src="img/nofollow.gif" width="15" height="15"><font size="2" color="#000000">ͯ��������̳&gt;&gt;&gt;��½��̳&gt;&gt;&gt; 
      </font></font></td>
  </tr>
</table>
<form name="form1" method="post" action="loginok.asp">
  <table width="302" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" cellspacing="0" cellpadding="0" align="center" height="129">
    <tr> 
      <td> 
        <div align="center"> 
          <p align="left"><font size="2">
            �û����ƣ� 
            <input type="text" name="user" size="19" STYLE="background-color:transparent;border:1px solid #000000">
            <br>
            <br>
            �û���� 
            <input type="password" name="pwd" size="20" STYLE="background-color:transparent;border:1px solid #000000">
            </font>
          <p align="left"><font size="2"> 
            <input type="submit" name="Submit" value="��½" STYLE="background-color:transparent;border:1px solid #000000">
            <input type="reset" name="Submit2" value="ȡ��" STYLE="background-color:transparent;border:1px solid #000000">
            [<a href="reg.asp">û��ע��</a>][<a href="repwd.asp">��������</a>]</font></p>
        </div>
      </td>
    </tr>
  </table>
</form>
<br>
<table width="755" border="0" cellspacing="0" cellpadding="0" align="center">
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
</body>
</html>
