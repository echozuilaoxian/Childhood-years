<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%
dim rs4 
dim sql4 
set rs4 = server.createobject("adodb.recordset") 
sql4 = "select*from wenzhang where id="&Trim(Request.form("id"))  
rs4.Open sql4,conn,1
if rs4("lock")=1 then
response.write"<script>alert('�Բ��𣬱������Ѿ���������ֹ�ظ���');history.back(-1);</script>"
response.End()
end if
rs4.close %>
<% if session("user")="" then
rname=Replace(Request.Form("name"),"'","''")
pwd=md5(Replace(Request.Form("pwd"),"'","''"))
sql="select * from user where user='"&rname&"' and pwd='"&pwd&"'"
 set rs=server.CreateObject("adodb.recordset")
 rs.open sql,conn,1,3
 if rs.eof then
 response.Write"<script> alert('�Բ����û������������');history.back(-1);</script>"
 response.End()
 end if%>
 <% end if %>
 <html><head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ɹ�</title>
<STYLE type=text/css>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 12pt
}
A:link {
	COLOR: #0080ff; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0080ff; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ee9c00; TEXT-DECORATION: underline
}
</STYLE>
<meta http-equiv=refresh content=3;url=default.asp></head>
<body bgcolor="#145EAF" text="#FFFFFF" link="#FF0000" vlink="#FF0000" alink="#FF0000" background="img/background.jpg"> 
<div align="center">
  <center>
    <table border="0" cellspacing="0" style="border-collapse: collapse" width="95%" bordercolorlight="#000000" cellpadding="0" bordercolordark="#FFFFFF">
      <tr> 
        <td width="100%" align="center"> <font color="#000000"> 
          <% 
		  rname=Replace(Request.Form("name"),"'","''")
title=Replace(Request.Form("title"),"'","''")
rcontent=Replace(Request.Form("GuestGuestContent"),"'","''")
rid=Replace(Request.Form("id"),"'","''")
rip=Replace(Request.Form("ip"),"'","''")
%>
          <% if rname="" or  rcontent="" then %>
          ��<a href="javascript:history.go(-1)">����</a>��д�������ϣ�����ܷ������ӣ� 
          <%else%>
          <BR>
          <% set savebbs=conn.execute("insert into rwenzhang(rname,rcontent,rid,rip)values('"&rname&"','"&rcontent&"','"&rid&"','"&rip&"')")%>
          </font> <font color="#000000">&nbsp; ��ӳɹ���<a href="show.asp?id=<%=request.form("id")%>">����</a>��<a href="default.asp">����</a> 
          <% conn.execute("update user set GDP=GDP+10 where user='"&rname&"'") %>
          <%end if   
set savebbs=nothing 
%>
          </font></td>
      </tr>
    </table>
  </center>
</div>
</body>
</html>