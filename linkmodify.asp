<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>友情连接修改系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
dim rs 
dim sql 
set rs = server.createobject("adodb.recordset") 
sql = "select*from link where id="&request.querystring("id")
rs.Open sql,conn,1,3
%>
<br>
<form name="form1" method="post" action="linkmodifyok.asp">
  <table width="414" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="77"><font size="2">站点编号：</font></td>
      <td width="337"><font size="2"> 
        <input type="text" name="id" value="<%=rs("id")%>" size="40">
        </font></td>
    </tr>
    <tr> 
      <td width="77"><font size="2">站点名称：</font></td>
      <td width="337"><font size="2"> 
        <input type="text" name="sitename" value="<%=rs("sitename")%>" size="40">
        </font></td>
    </tr>
    <tr> 
      <td width="77"><font size="2">站点地址：</font></td>
      <td width="337"><font size="2"> 
        <input type="text" name="siteurl" value="<%=rs("siteurl")%>" size="40">
        </font></td>
    </tr>
    <tr> 
      <td width="77"><font size="2">logo地址：</font></td>
      <td width="337"><font size="2"> 
        <input type="text" name="logourl" value="<%=rs("logourl")%>" size="40">
        </font></td>
    </tr>
    <tr> 
      <td width="77">&nbsp;</td>
      <td width="337"> 
        <input type="submit" name="Submit" value="修改">
        <input type="reset" name="Submit2" value="Reset">
      </td>
    </tr>
  </table>
</form>
<%
rs.close 
%>
<br>
</body>
</html>
