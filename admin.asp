<%@ CODEPAGE ="936"%>
<html>
<% if session("user")="" then
response.write"<script>alert('对不起，您没有权限访问此页面，如果您是管理员，请先登陆！<br>如果您不是管理员，请与管理员联系！');history.back(-1);</script>"
response.End()
end if %> 
<head>
<title>论坛管理中心--纯真论坛</title>
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
