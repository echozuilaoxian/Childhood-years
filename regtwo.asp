<%@ CODEPAGE ="936"%>
<html>
<head>
<title>��������дע����Ϣ--ͯ��������̳��</title>
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
</STYLE><SCRIPT LANUGAGE="JavaScript">
<!--
function pop(newpage) {
var
popwin=window.open(newpage,"popWin","scrollbars=no,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,width=250,height=450");
return false;
}
//-->
</SCRIPT></head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="10">
<table width="755" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td><img src="img/banner.jpg" width="760" height="80" border="1"></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="27" bgcolor="#FFFFFF"><font size="2" color="#000000"><font color="#FF0000"><%=session("user")%><font color="#000000"> 
      </font></font><img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="default.asp">��ҳ</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <a href="reg.asp">ע��</a> <img src="img/splt.gif" width="7" height="15" align="absmiddle"> 
      <% if session("user")="" then response.write"<font size=2><a href=login.asp>��½</a></font>" else response.write"<font size=2><a href=out.asp?user="&user&">�˳�</a></font>" end if %>
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="search.asp">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">����Ա���</a> 
      </font></td>
  </tr>
</table>
<br>
<table width="417" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" cellspacing="0" cellpadding="0" align="center" height="27">
  <tr> 
    <td bgcolor="#F5F5F5" height="24"><font color="#000000">��ҳ &gt;&gt; ���û�ע�� 
      </font></td>
  </tr>
</table>
<form name="form1" method="post" action="regok.asp">
  <table width="416" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" cellspacing="0" cellpadding="0" align="center" height="485">
    <tr bgcolor="#F5F5F5"> 
      <td colspan="2" height="22"> 
        <div align="center"><font size="2" color="#000000">::����ȷ��д�й���Ϣ����<font color="#FF0000">*</font>��Ϊ������)::</font></div>
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right"><font size="2">�û����ƣ�</font></div>
      </td>
      <td width="208"> 
        <div align="left"><font size="2" color="#FF0000"> * 
          <input type="text" name="user" size="19" maxlength="8" STYLE="background-color:transparent;border:1px solid #000000">
          </font></div>
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right"><font size="2">�û����</font></div>
      </td>
      <td width="208"> 
        <div align="left"><font size="2" color="#FF0000"> * 
          <input type="password" name="pwd" size="20"STYLE="background-color:transparent;border:1px solid #000000" >
          </font></div>
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right">�ظ����</div>
      </td>
      <td width="208"> <font color="#FF0000">*</font> 
        <input type="password" name="textfield" size="20" STYLE="background-color:transparent;border:1px solid #000000">
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right"><font size="2">�� �� ��</font></div>
      </td>
      <td width="208"> 
        <div align="left"><font size="2" color="#FF0000"> * 
          <select name="sex">
            <option value="˧��">˧��</option>
            <option value="��Ů">��Ů</option>
          </select>
          </font></div>
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right"><font size="2">OICQ��</font></div>
      </td>
      <td width="208"> 
        <div align="left"><font size="2" color="#FF0000"> * 
          <input type="text" name="QQ" size="20" STYLE="background-color:transparent;border:1px solid #000000">
          </font></div>
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right"><font size="2">E-mail:</font></div>
      </td>
      <td width="208"> 
        <div align="left"><font size="2" color="#FF0000"> * 
          <input type="text" name="email" size="25" STYLE="background-color:transparent;border:1px solid #000000">
          </font></div>
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right">������ʾ���⣺</div>
      </td>
      <td width="208"><font color="#FF0000">*</font> 
        <input type="text" name="mimatishi" size="25" STYLE="background-color:transparent;border:1px solid #000000">
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right">������ʾ�𰸣�</div>
      </td>
      <td width="208"><font color="#FF0000">*</font> 
        <input type="text" name="tishidaan" size="25" STYLE="background-color:transparent;border:1px solid #000000">
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right">������ҳ��</div>
      </td>
      <td width="208"><font color="#FFFFFF">*</font> 
        <input type="text" name="homepage" size="25" STYLE="background-color:transparent;border:1px solid #000000">
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right">��ϵ��ַ��</div>
      </td>
      <td width="208"><font color="#FFFFFF">* </font> 
        <input type="text" name="address" size="25" STYLE="background-color:transparent;border:1px solid #000000">
      </td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right">�û�ͷ��</div>
      </td>
      <td width="208"> <font color="#FFFFFF">* </font> 
        <input type="text" name="userico" size="5"STYLE="background-color:transparent;border:1px solid #000000" >
        <A href="ico.htm" onclick="return pop(this.href);">ͷ�����</a></td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right">����ǩ����</div>
      </td>
      <td rowspan="2" width="208"> <font color="#FFFFFF">* </font> 
        <textarea name="sign" cols="25" rows="4" STYLE="background-color:transparent;border:1px solid #000000"></textarea>
      </td>
    </tr>
    <tr> 
      <td width="176">&nbsp;</td>
    </tr>
    <tr> 
      <td width="176"> 
        <div align="right"><font size="2"> 
          <input type="submit" name="Submit" value="ע��" STYLE="background-color:transparent;border:1px solid #000000">
          </font></div>
      </td>
      <td width="208"> 
        <div align="left"><font size="2"> 
          <input type="reset" name="Submit2" value="����" STYLE="background-color:transparent;border:1px solid #000000">
          </font></div>
      </td>
    </tr>
  </table>
</form>
<br>
<table width="775" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td bgcolor="#F5F5F5" height="24"> 
      <div align="center"><font color="#000000">| ��վ��� | �������� | �������� | QQ��47598622 
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
