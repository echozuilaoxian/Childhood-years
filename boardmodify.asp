<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>修改成功！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<% dim rs
dim sql
set rs=server.createobject("adodb.recordset")
sql="select*from board where id="&trim(request.querystring("id"))
on error resume next 
rs.Open sql,conn,1
RS.MoveFirst
do while not rs.eof
%>
<form name="form1" method="post" action="chenggong.asp">
  <table width="517" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="74" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">版块：</font></div>
      </td>
      <td width="443"><font size="2"> 
        <input type="text" name="boardname" value="<%=rs("boardname")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="74" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">图标：</font></div>
      </td>
      <td width="443"> 
        <input type="text" name="ico" value="<%=rs("ico")%>">
      </td>
    </tr>
    <tr> 
      <td width="74" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">斑竹：</font></div>
      </td>
      <td width="443"> <font size="2"> 
        <input type="text" name="banzhu" value="<%=rs("banzhu")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="74" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">简介：</font></div>
      </td>
      <td width="443"><font size="2"> 
        <textarea name="boardcontent" cols="35" rows="5"><%=rs("boardcontent")%></textarea>
        </font></td>
    </tr>
    <tr> 
      <td width="74" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2" color="#000000">大类：</font></div>
      </td>
      <td width="443"> 
        <input type="text" name="typeid" value="<%=rs("typeid")%>">
      </td>
    </tr>
    <tr> 
      <td width="74" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2">ID: </font></div>
      </td>
      <td width="443"> <font size="2"> 
        <input type="text" name="id" value="<%=rs("id")%>">
        </font></td>
    </tr>
    <tr>
      <td width="74" bgcolor="#F5F5F5"> 
        <div align="right"><font size="2">属性：</font></div>
      </td>
      <td width="443">
        <input type="text" name="lock" value="<%=rs("lock")%>" size="10">
        <font color="#FF0000" size="2"> 注意：&quot;1&quot;为版面锁定，&quot;0&quot;为版面正常</font></td>
    </tr>
    <tr> 
      <td width="74"><font size="2"></font></td>
      <td width="443"> <font size="2"> 
        <input type="submit" name="Submit" value="修改">
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

