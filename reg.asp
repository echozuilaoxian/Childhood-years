<%@ CODEPAGE ="936"%>
<html>
<head>
<title>�û�ע��--ͯ��������̳��</title>
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
      <img src="img/splt.gif" width="7" height="15" align="absmiddle"> <a href="guest/index.asp" target="_blank">����</a> 
      <img src="img/splt.gif" width="7" height="15" align="absbottom"> <a href="adminlogin.asp">����Ա���</a> 
      </font></td>
  </tr>
</table>
<br>
<table width="471" border="0" bordercolorlight="#FFFFFF" bordercolordark="#003399" cellspacing="0" cellpadding="0" align="center" bordercolor="#003399">
  <tr> 
    <td bgcolor="#F5F5F5" height="24">��ҳ &gt;&gt;&gt; ���û�ע�� </td>
  </tr>
</table>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="467" align="center">
  <tr> 
    <td height="228"> �� <b>����ע��ǰ�����Ķ���̳Э�� </b> 
      <p> ��ӭ�����뱾վ��μӽ��������ۣ���վ��Ϊ������̳��Ϊά�����Ϲ������������ȶ��������Ծ������������</p>
      <p>һ���������ñ�վΣ�����Ұ�ȫ��й¶�������ܣ������ַ�������Ἧ��ĺ͹���ĺϷ�Ȩ�棬������<br>
        �ñ�վ���������ƺʹ���������Ϣ�� <br>
        ������һ��ɿ�����ܡ��ƻ��ܷ��ͷ��ɡ���������ʵʩ�ģ�<br>
        ����������ɿ���߸�������Ȩ���Ʒ���������ƶȵģ�<br>
        ����������ɿ�����ѹ��ҡ��ƻ�����ͳһ�ģ�<br>
        �������ģ�ɿ�������ޡ��������ӣ��ƻ������Ž�ģ�<br>
        �������壩�������������ʵ��ɢ��ҥ�ԣ������������ģ� <br>
        ��������������⽨���š����ࡢɫ�顢�Ĳ�����������ɱ���ֲ�����������ģ�<br>
        �������ߣ���Ȼ�������˻���������ʵ�̰����˵ģ����߽����������⹥���ģ�<br>
        �������ˣ��𺦹��һ��������ģ�<br>
        �������ţ�����Υ���ܷ��ͷ�����������ģ�<br>
        ������ʮ��������ҵ�����Ϊ�ġ�<br>
        �����������أ����Լ������ۺ���Ϊ����<br>
      </p>
      <p>�� [<a href="regtwo.asp">ͬ��</a>]
    </td>
  </tr>
</table>
<br>
<table width="775" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="24" bgcolor="#F5F5F5"> 
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
