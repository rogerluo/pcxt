<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	Set rs = server.CreateObject("Adodb.Recordset")
	strsql="select sysstate from [adminuser]  where adminname='"&session("adminname")&"'"
	Set rs = conn.execute(strsql)
	sysstate = rs("sysstate")
	If request.Form("state")<>"" Then 
		strsql="update [adminuser] set sysstate="&CInt(request.Form("state"))&" where adminname='"&session("adminname")&"'"
		conn.execute(strsql)
		session("sysstate") = CInt(request.Form("state"))
		rs.close
		Set rs=Nothing
		conn.close
		set conn = nothing
		response.write "<script>alert('设置修改成功！');window.location.href='adminmain.asp';</script>"
	Else
		rs.close
		Set rs=Nothing
		conn.close
		Set conn=nothing
	End If 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>电子信息工程学院研究生课程状况调查系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="systemstate" method="post" action="">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生课程状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>，你好！</td>
    <td align="right">【<a href="adminmain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="adminlogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td><li><a href="adminlist.asp">课程列表</a></td>
  </tr>
  <tr>
    <td><li><input type="submit" name="Submit" value="<%
		If sysstate = 1 Then
			response.write "关闭"
		Else
			response.write "开启"
		End If
	%>系统"><input name="state" type="hidden" id="state" value="<%
			If sysstate = 1 Then
				response.write "0" 
			Else
				response.write "1" 
			End If%>"></td>
  </tr>
  <tr>
    <td><li><a href="adminupload.asp">上传文件</td>
  </tr>
</table>
</form>
</body>
</html>
