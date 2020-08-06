<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>链接插入成功-请稍后……</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=2;url=linkadmin.asp>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<% sitename=Replace(Request.Form("sitename"),"'","''")
   siteurl=Replace(Request.Form("siteurl"),"'","''")
   logourl=Replace(Request.Form("logourl"),"'","''") %>
<% set save=conn.execute("insert into link(sitename,siteurl,logourl)values('"&sitename&"','"&siteurl&"','"&logourl&"')")%>
<% set save=nothing %>
<font size="2" color="#FF0000">新链接插入成功，请稍后…………</font> 
</body>
</html>
