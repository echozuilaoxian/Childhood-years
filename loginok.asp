<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<html><title>登陆成功--童真岁月论坛</title>
<head>
<meta http-equiv=refresh content=2;url=default.asp></head>
<body>
<%  if request.Form("user")="" or request.form("pwd")="" then
    response.Write"<script language=javascript>alert('请填写完整');history.back(-1);</script>"
 response.End()
 end if
    user=Replace(Request.Form("user"),"'","''")
    pwd=md5(Replace(Request.Form("pwd"),"'","''"))
 sql="select * from user where user='"&user&"' and pwd='"&pwd&"'"
 set rs=server.CreateObject("adodb.recordset")
 rs.open sql,conn,1,3
 if rs.eof then
 response.Write"<script> alert('用户名或密码错误');history.back(-1);</script>"
 response.End()
 else  
    session("user")=rs("user")
    session("pwd")=rs("pwd")
    response.write"<font size=2>欢迎您的光临，</font>"%>
<font size=2 color=#FF0000><%=session("user")%></font> 
<% conn.execute("update user set GDP=GDP+5 where user='"&user&"'") %>
<% set savebbs=nothing %>
<% user=session("user")
   userid=session.sessionid %>
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select*from online where user='"&user&"'"
rs.Open sql,conn,1,3 
if rs.eof then
%>
<% set save=conn.execute("insert into online(user,userid)values('"&user&"','"&userid&"')")%>
<% set save=nothing %>
<% end if %>
<%end if%>
</body>
</html>



