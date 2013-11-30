<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>电子信息工程学院研究生教学状况调查系统</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="stuevaluate" method="post" action="stuchkevaluate.asp">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生教学状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname")%>，你好！</td>
    <td align="right">【<a href="stumain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="stulogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
<%
	strsql = "select [courseinfo].coursedate,[courseinfo].coursename,[courseinfo].courseteacher,[courseinfo].coursetype,[courseinfo].courseid,[coursemember].courseorder,[coursemember].isevaluate from [coursemember] join [courseinfo] on ([courseinfo].courseid = [coursemember].courseid and [coursemember].stuid='"+session("stuid")+"' and [coursemember].isevaluate = 1 and [courseinfo].courseorder=[coursemember].courseorder)"
	Set rs = conn.execute(strsql)
	If Not (rs.bof And rs.eof) Then
%>
	<table width="550"  border="1" align="center" cellpadding="5" cellspacing="0" >
		<tr><td align="center" class="titlebg">已修课程列表</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr align="center">
						<td width=200>课程名称</td>
						<td width=80>授课老师</td>
						<td width=80>课程类型</td>
						<td width=80>是否已评价</td>
						<td>编辑</td>
					</tr>
					
<%
	Do While Not rs.eof
		response.write "<tr align='center'><td>" + rs("coursename") + "</td><td>" + rs("courseteacher") + "</td><td>" + rs("coursetype") + "</td><td>是</td><td><a href='stulistdetail.asp?courseid="+rs("courseid")+"&coursedate="+cstr(rs("coursedate"))+"&courseorder="+cstr(rs("courseorder"))+"'>编辑</a></td></tr>"
		
		rs.movenext
	Loop 
%>
			
				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>
<%
	Else 
%>
	<table width="550"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td  class="titlebg" bgcolor='#cccccc'>没有任何课程进行评价</td></tr>
	</table>

</form>
</body>
</html>
<%
	End If 
%>
