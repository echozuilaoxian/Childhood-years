<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<html><title>�����½�ɹ�--ͯ��������̳</title>
<head>
</head>
<body>
<%  if request.Form("user")="" or request.form("pwd")="" then
    response.Write"<script language=javascript>alert('����д����');history.back(-1);</script>"
 response.End()
 end if
    user=Replace(Request.Form("user"),"'","''")
    pwd=md5(Replace(Request.Form("pwd"),"'","''"))
 sql="select * from admin where user='"&user&"' and pwd='"&pwd&"'"
 set rs=server.CreateObject("adodb.recordset")
 rs.open sql,conn,1,1
 if rs.eof then
 response.Write"<script> alert('�û������������');history.back(-1);</script>"
 response.End()
 else  
    session("user")=rs("user")
    session("pwd")=rs("pwd")
    response.write"��ӭ���Ĺ��٣�"%>
<%=session("user")%> 
<% set savebbs=nothing %>
<%response.redirect"admin.asp"%>
<%end if%>
</body>
</html>



