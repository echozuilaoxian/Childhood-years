<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head>
<title>看版添加成功！请稍后……</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content=2;url=boardadmin.asp></head>

<body bgcolor="#FFFFFF" text="#000000">
<% boardname=Replace(Request.Form("boardname"),"'","''")
ico=Replace(Request.Form("ico"),"'","''")
banzhu=Replace(Request.Form("banzhu"),"'","''")
boardcontent=Replace(Request.Form("boardcontent"),"'","''")
typeid=Replace(Request.Form("typeid"),"'","''")
%>
<% if boardname="" or banzhu="" or boardcontent="" or typeid="" then
response.write"<script>alert('对不起，请填写完整');history.back(-1);</script>"
response.End()
else 
conn.execute("insert into board(boardname,ico,banzhu,boardcontent,typeid)values('"&boardname&"','"&ico&"','"&banzhu&"','"&boardcontent&"','"&typeid&"')")
set conn=nothing
end if %>
<font size="2" color="#FF0000">看版添加成功！请稍后…………</font> 
</body>
</html>
