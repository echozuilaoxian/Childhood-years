<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"--><html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>发表成功</title>
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
          <meta http-equiv=refresh content=100;url=default.asp>
</head>
<body bgcolor="#145EAF" text="#FFFFFF" link="#FF0000" vlink="#FF0000" alink="#FF0000" background="img/background.jpg">
<div align="center"><center>
    <table border="0" cellspacing="0" style="border-collapse: collapse" width="95%" bordercolorlight="#000000" cellpadding="0" bordercolordark="#FFFFFF">
      <tr>
        <td width="100%" align="center"> <font color="#000000">
          <% name=Replace(Request.Form("name"),"'","''")
title=Replace(Request.Form("title"),"'","''")
pwd=Replace(Request.Form("pwd"),"'","''")
biaoqing=Replace(Request.Form("biaoqing"),"'","''")
content=Replace(Request.Form("GuestContent"),"'","''")
boardid=Replace(Request.Form("boardid"),"'","''")
ip=Replace(Request.Form("ip"),"'","''")
%>
          <% if name="" or title="" or pwd="" or content="" then %>
          请<a href="javascript:history.go(-1)">后退</a>填写完整资料，你才能发表帖子！ 
          <%else%>
          <BR>
            <% set savebbs=conn.execute("insert into wenzhang(name,title,biaoqing,content,boardid,ip)values('"&name&"','"&title&"','"&biaoqing&"','"&content&"','"&boardid&"','"&ip&"')")%>
          </font>
          <font color="#000000">&nbsp; 添加成功！·<a href="board.asp?id=<%=request.form("boardid")%>">返回</a>
          <% conn.execute("update user set GDP=GDP+15 where user='"&name&"'") %>
		  <% end if %>
		  <%  
set savebbs=nothing 
%> 
          </font></td> 
      </tr></table></center></div></body></html>