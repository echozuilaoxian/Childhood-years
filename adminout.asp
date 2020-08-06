<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html>
<head><title><% name=session("user") %>谢谢您！您已经成功退出了论坛！</title></head><meta http-equiv=refresh content=3;url=default.asp>
<body background="img/background.jpg">
<font size="2"> 
<% user=session("user") %>
<% conn.execute("update user set online=0 where user='"&user&"'")%>
<% set del=nothing %>

</font> 
<div align="center"><font size="2" color="#FF0000"><%=session("user")%>谢谢您，您已经成功退出了管理！</font><font size="2"> 
  <font color="#FF0000">本页将自动转向，请等待…………</font></font></div>
<% Session.Abandon
session("user")="" %></body>
</html>