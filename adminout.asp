<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head><title><% name=session("user") %>лл�������Ѿ��ɹ��˳�����̳��</title></head><meta http-equiv=refresh content=3;url=default.asp>
<body background="img/background.jpg">
<font size="2"> 
<% user=session("user") %>
<% conn.execute("update user set online=0 where user='"&user&"'")%>
<% set del=nothing %>

</font> 
<div align="center"><font size="2" color="#FF0000"><%=session("user")%>лл�������Ѿ��ɹ��˳��˹���</font><font size="2"> 
  <font color="#FF0000">��ҳ���Զ�ת����ȴ���������</font></font></div>
<% Session.Abandon
session("user")="" %></body>
</html>