<%@ CODEPAGE ="936"%>
<html>
<% if session("user")="" then
response.write"<script>alert('�Բ�����û��Ȩ�޷��ʴ�ҳ�棬������ǹ���Ա�����ȵ�½��<br>��������ǹ���Ա���������Ա��ϵ��');history.back(-1);</script>"
response.End()
end if %> 
<head>
<title>��̳��������--������̳</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<frameset cols="174,*" frameborder="NO" border="1" framespacing="1" rows="*" bordercolor="#FF0000"> 
  <frame name="leftFrame" scrolling="NO" noresize src="admin4.asp">
  <frame name="mainFrame" src="admin2.asp">
</frameset>
<noframes>
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
