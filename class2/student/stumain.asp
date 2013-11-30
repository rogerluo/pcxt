<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<%
	strsql = "select stuid from evaluate2_gra where stuid = '"+session("stuid_gra")+"'"
	Set rs = conn.execute(strsql)
	If rs.eof Or rs.bof Then 
		isevaluate = 0
	Else 
		isevaluate = 1
	End If 
	rs.close
	Set rs = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>电子信息工程学院研究生课程状况调查系统</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生课程状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname_gra")%>，你好！</td>
    <td align="right">【<a href="stumain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="stulogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td>
		<li><a href="stutotalevaluate.asp">课程总体评价</a>
<%
	If isevaluate = 0 Then 
		response.write "(尚没评价)"
	Else 
		response.write "(已评价)"
	End If 
%>	
	</td>
  </tr>	
  <tr>
    <td><li><a href="stuevaluate.asp">尚没评价课程</a></td>
  </tr>
  <tr>
    <td><li><a href="stulist.asp">查看已评课程</td>
  </tr>
</table>
</body>
</html>