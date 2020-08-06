<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<html><title>管理登陆成功--童真岁月论坛</title>
<head>
</head>
<body>
<%  if request.Form("user")="" or request.form("pwd")="" then
    response.Write"<script language=javascript>alert('请填写完整');history.back(-1);</script>"
 response.End()
 end if
    user=Replace(Request.Form("user"),"'","''")
    pwd=md5(Replace(Request.Form("pwd"),"'","''"))
 sql="select * from admin where user='"&user&"' and pwd='"&pwd&"'"
 set rs=server.CreateObject("adodb.recordset")
 rs.open sql,conn,1,1
 if rs.eof then
 response.Write"<script> alert('用户名或密码错误');history.back(-1);</script>"
 response.End()
 else  
    session("user")=rs("user")
    session("pwd")=rs("pwd")
    response.write"欢迎您的光临，"%>
<%=session("user")%> 
<% set savebbs=nothing %>
<%response.redirect"admin.asp"%>
<%end if%>
</body>
</html>



