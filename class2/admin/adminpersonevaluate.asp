<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	'获取链接参数
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))	
	stuid = Trim(request.querystring("stuid"))

	If courseid = "" Or Not IsNumeric(courseid) Or stuid = "" Or Not IsNumeric(stuid) Or coursedate = "" Then 
		conn.close()
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');window.location.href='stuevaluate.asp';</script>"
		response.end
	End If 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>电子信息工程学院研究生课程状况调查系统</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>
<%
	strsql = "select * from [courseinfo_gra] where courseid='"+courseid+"' and coursedate ='"+coursedate+"'"
	Set rs = conn.execute(strsql)

	If rs.eof Or rs.bof Then
		rs.close
		Set rs = Nothing 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('没有相应课程信息！');window.location.href='adminlist.asp';</script>"
		response.end
	End If

	coursename = CStr(rs("coursename"))
	coursetime = CStr(rs("coursetime"))
	coursecredit = cstr(rs("coursecredit"))
	courseteacher = cstr(rs("courseteacher"))
	coursetype = cstr(rs("coursetype"))
	coursetype2 = cstr(rs("coursetype2"))

	courseheadcount = cstr(rs("courseheadcount"))

	rs.close

	strsql = "select * from [coursemember_gra] where courseid='"+courseid+"' and stuid ='"+stuid+"' and coursedate ='"+coursedate+"'"
	Set rs = conn.execute(strsql)
	If rs.eof Or rs.bof Then 
		rs.close
		Set rs = Nothing 
		conn.close
		Set conn = Nothing
		response.write "<script>alert('没有相应学生信息！');window.location.href='adminlist.asp';</script>"
	End If 

	stuname = rs("stuname")
	stuobject = rs("stuobject")
	stucollege = rs("stucollege")
	rs.close

	strsql = "select * from [evaluate_gra] where courseid='"+courseid+"' and stuid ='"+stuid+"' and coursedate='"+coursedate+"'"
	Set rs = conn.execute(strsql)

	coursecontent = rs("coursecontent")
	coursedifficult = rs("coursedifficult")
	coursepoint = rs("coursepoint")
	courselang = rs("courselang")
	courseknowhelp = rs("courseknowhelp")
	coursejobhelp = rs("coursejobhelp")
	coursecontinue = rs("coursecontinue")

	rs.close
%>
<body>

<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生课程状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>，你好！</td>
    <td align="right">【<a href="adminlistdetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>">返回上页</a>】&nbsp;&nbsp;【<a href="adminmain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="adminlogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="3" cellspacing="0">
		<tr align="center" class="titlebg"><td>课程基本信息</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width="50%">课程名&nbsp;&nbsp;：<%=coursename%></td>
						<td width="25%">授课教师：<%=courseteacher%></td>
						<td>课程类别：<%=coursetype%></td>
					</tr>
					<tr>
						<td>授课人数：<%=courseheadcount%></td>
						<td>类别&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;：<%=coursetype2%></td>
						<td>学分&nbsp;&nbsp;：<%=coursecredit%> 学时&nbsp;&nbsp;：<%=coursetime%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td>&nbsp;&nbsp;</td></tr>
		<tr align="center" class="titlebg"><td>学生基本信息</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width="50%">姓名&nbsp;&nbsp;：<%=stuname%></td>
						<td>学号&nbsp;&nbsp;：<%=stuid%></td>
						
					</tr>
					<tr>
						<td width=400>专业&nbsp;&nbsp;：<%=stuobject%></td>
						<td width>院系&nbsp;&nbsp;：<%=stucollege%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td>&nbsp;&nbsp;</td></tr>
		<tr align="center" class="titlebg"><td>学生评价信息</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td>评价内容（<font color=red>红色为主要考查项</font>）</td>
						<td>评价选项（带*为必须填写）</td>
					</tr>
					<tr>
					<td width=50%>课程内容偏重于</td>
					<td>
						<%
							Select Case coursecontent
								Case "A"
								response.write "A.科研理论"
								Case "B"
								response.write "B.具体实践"
								Case "C"
								response.write "C.两者兼顾"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>课程学习难度</td>
					<td>
						<%
							Select Case coursedifficult
								Case "A"
								response.write "A.容易"
								Case "B"
								response.write "B.一般"
								Case "C"
								response.write "C.比较难"
								Case "D"
								response.write "D.很难"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>课程内容知识点</td>
					<td>
						<%
							Select Case coursepoint
								Case "A"
								response.write "A.专业相关性大"
								Case "B"
								response.write "B.就业相关性大"
								Case "C"
								response.write "C.两者均否"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>课程语种建议</td>
					<td>
						<%
							Select Case courselang
								Case "A"
								response.write "A.中文"
								Case "B"
								response.write "B.英文"
								Case "C"
								response.write "C.中英结合"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>课程对自身知识面帮助</font></td>
					<td>
						<%
							Select Case courseknowhelp
								Case "A"
								response.write "A.很大"
								Case "B"
								response.write "B.一般"
								Case "C"
								response.write "C.很少"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>课程对自身找工作帮助</font></td>
					<td>
						<%
							Select Case coursejobhelp
								Case "A"
								response.write "A.很大"
								Case "B"
								response.write "B.一般"
								Case "C"
								response.write "C.很少"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>课程是否继续</font></td>
					<td>
						<%
							Select Case coursecontinue
								Case "A"
								response.write "A.同意"
								Case "B"
								response.write "B.列入考察范围"
								Case "C"
								response.write "C.建议更换别的课程"
							End Select 
						%>
					</td>
				</tr>
				</table>
			</td>
		</tr>
	</table>

</body>
</html>