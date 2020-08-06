<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>删除帖子成功！-- 童真岁月论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from wenzhang where id="&trim(request.querystring("id"))
on error resume next 
rs.Open sql,conn,1
RS.MoveFirst
do while not rs.eof
%>
<form name="form1" method="post" action="showdelok.asp" onsubmit="return confirm('此操作将不可恢复，确定删除？')">
  <table width="517" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="84"><font size="2">看版编号：</font></td>
      <td width="433"> 
        <input type="text" name="boardid" value="<%=rs("boardid")%>">
      </td>
    </tr>
    <tr> 
      <td width="84"><font size="2">帖子标题：</font></td>
      <td width="433"> <font size="2"> 
        <input type="text" name="title" value="<%=rs("title")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="84"><font size="2">帖子内容：</font></td>
      <td width="433"><font size="2"> 
        <textarea name="content" cols="35" rows="5"><%=rs("content")%></textarea>
        </font></td>
    </tr>
    <tr> 
      <td width="84"><font size="2">帖子编号</font>：</td>
      <td width="433"> <font size="2"> 
        <input type="text" name="id" value="<%=rs("id")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="84"><font size="2"></font></td>
      <td width="433"> <font size="2"> 
        <input type="submit" name="Submit" value="删除">
        <input type="reset" name="Submit2" value="Reset">
        </font></td>
    </tr>
  </table>
</form>
<% 
rs.movenext  
loop  
rs.Close  
%></body>
</html>

