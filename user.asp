<!--#include file="conn.asp"-->
<% response.expires="-100" %>
<% dim rsc
dim sqlc
set rsc=server.createobject("adodb.recordset")
sqlc="select* from online where ltime<now()-0.005"
rsc.open sqlc,conn,1,3
if rsc.eof then
else
do while not rsc.eof 
rsc.delete
rsc.movenext
loop
rsc.close
end if
%>
<% set co=server.createobject("adodb.recordset")
	co.open "select count(id) as sum from online",conn,1,3
	if not co.eof then %>
<%
strsql="select * from online"
set rs=server.createobject("ADODB.recordset")
rs.open strsql,conn,2,1
if rs.eof then
response.write("<font size=2><font color=#FF0000>现在没有注册用户在线!</font></font>")
end if
%>
<html>
<head>
<title>在线详细列表情况--学习论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="img/line.gif">
<p><font size="2">[<a href="javascript:location.reload()">刷新</a>]</font></p>
<p><font size="2"> </font> </p>
<table width="600" align="center" border="0" height="8">
  <%for i=1 to co("sum")
if (i mod 5=1) then
response.write"<tr>"
end if
response.write"<td><img src=img/online_member.gif><font size=2>"&rs("user")&"</font></td>"
if (i mod 5=0) then
response.write"<tr></tr>"
end if
rs.movenext
next%>
</table>
<font size="2"> 
<%
	end if
	co.close
	set co=nothing
	%>
</font> 
</body>
</html>

