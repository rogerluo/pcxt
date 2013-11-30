<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	courseorder = Trim(request.querystring("courseorder"))
	stuid = Trim(request.querystring("stuid"))

	If courseid = "" Or Not IsNumeric(courseid) Or stuid = "" or coursedate = "" or courseorder = "" or Not IsNumeric(courseorder) Then 
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
<title>电子信息工程学院研究生教学状况调查系统</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>

<%
	strsql = "select * from [courseinfo] where courseid='"+courseid+"' and coursedate ='"+coursedate+"' and courseorder = "+ courseorder
	Set rs = conn.execute(strsql)
	If rs.eof Or rs.bof Then
		rs.close
		Set rs = Nothing 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('没有相应课程信息！');window.location.href='adminlist.asp';</script>"
		response.end
	End If 
	coursename = rs("coursename")
	coursetime = CStr(rs("coursetime"))
	courseteacher = cstr(rs("courseteacher"))
	coursetype = cstr(rs("coursetype"))
	courseheadcount = cstr(rs("courseheadcount"))
	coursecredit = cstr(rs("coursecredit"))
	courseteacher2 = cstr(rs("courseteacher2"))
	rs.close

	strsql = "select * from [coursemember] where courseid='"+courseid+"' and stuid ='"+stuid+"' and coursedate ='"+coursedate+"' and courseorder="+courseorder
	Set rs = conn.execute(strsql)
	stuname = rs("stuname")
	stuobject = rs("stuobject")
	stucollege = rs("stucollege")
	rs.close

	strsql = "select * from [evaluate] where courseid='"+courseid+"' and stuid ='"+stuid+"' and coursedate='"+coursedate+"' and courseorder="+courseorder
	Set rs = conn.execute(strsql)

	book = rs("book")
	courseraw = rs("courseraw")
	rawlang = rs("rawlang")
	lang = rs("lang")
	scale = rs("scale")
	spirit = rs("spirit")
	info = rs("info")
	teachtime = rs("teachtime")
	teachtime2 = rs("teachtime2")
	teachstatus = rs("teachstatus")
	teachername2 = rs("teachername2")
	stoptimes = rs("stoptimes")
	stopcourse = rs("stopcourse")
	retimes = rs("retimes")
	total = rs("total")
	advice = rs("advice")

	rs.close
%>
<body>

<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生教学状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>，你好！</td>
    <td align="right">【<a href="adminlistdetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>&target=0">返回上页</a>】&nbsp;&nbsp;【<a href="adminmain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="stulogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="3" cellspacing="0">
		<tr align="center" class="titlebg"><td bgcolor='#cccccc'>课程基本信息</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width="50%">课程名&nbsp;&nbsp;：<%=coursename%></td>
						<td width="25%">授课教师：<%=courseteacher%></td>
						<td>授课类型：<%=coursetype%></td>
					</tr>
					<tr>
						<td>上课地点：<%=coursewhere%></td>
						<td>授课系&nbsp;&nbsp;：<%=coursehost%></td>
						<td>学时&nbsp;&nbsp;&nbsp;&nbsp;：<%=coursetime%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td>&nbsp;&nbsp;</td></tr>
		<tr align="center" class="titlebg"><td bgcolor='#cccccc'>学生基本信息</td></tr>
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
		<tr align="center" class="titlebg"><td bgcolor='#cccccc'>学生评价信息</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td>评价内容（<font color=red>红色为主要考查项</font>）</td>
						<td>评价选项（带*为必须填写）</td>
					</tr>
					<tr>
					<td width=50%">使用教材</td>
					<td>
						<%
							Select Case book
								Case "A"
								response.write "A.正式出版教材"
								Case "B"
								response.write "B.自编讲义"
								Case "C"
								response.write "C.英文影印版"
								Case "D"
								response.write "D.英文原版教材"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>使用课件</td>
					<td>
						<%
							Select Case courseraw
								Case "A"
								response.write "A.正式出版课件"
								Case "B"
								response.write "B.自编课件"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>电子教案语种</font></td>
					<td>
						<%
							Select Case rawlang
								Case "A"
								response.write "A.英文"
								Case "B"
								response.write "B.中文"
								Case "C"
								response.write "C.中/英双语"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>授课语种</font></td>
					<td>
						<%
							Select Case lang
								Case "A"
								response.write "A.英文"
								Case "B"
								response.write "B.中文"
								Case "C"
								response.write "C.中/英双语"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>英语授课的比例</font></td>
					<td>
						<%
							Select Case scale
								Case "A"
								response.write "A.英语讲述70%以上"
								Case "B"
								response.write "B.英语讲述50%~70%"
								Case "C"
								response.write "C.英语讲述30%~50%"
								Case "D"
								response.write "D.英语讲述30%以下"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>准备充分，上课精神饱满</td>
					<td>
						<%
							Select Case spirit
								Case "A"
								response.write "A.很好"
								Case "B"
								response.write "B.比较好"
								Case "C"
								response.write "C.尚可"
								Case "D"
								response.write "D.不太好"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>教学内容充实具有前瞻性，提供足够的信息量</td>
					<td>
						<%
							Select Case info
								Case "A"
								response.write "A.很好"
								Case "B"
								response.write "B.比较好"
								Case "C"
								response.write "C.尚可"
								Case "D"
								response.write "D.不太好"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>对教学内容熟悉，讲述清楚</td>
					<td>
						<%
							Select Case teachstatus
								Case "A"
								response.write "A.很好"
								Case "B"
								response.write "B.比较好"
								Case "C"
								response.write "C.尚可"
								Case "D"
								response.write "D.不太好"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>开课教师本人授课学时</font></td>
					<td>
						<%
							Select Case teachtime
								Case "A"
								response.write "A.80%以上"
								Case "B"
								response.write "B.60%~80%"
								Case "C"
								response.write "C.40%~60%"
								Case "D"
								response.write "D.40%以下"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>替课教师授课学时</font></td>
					<td>
						<%
							Select Case teachtime2
								Case "A"
								response.write "A.80%以上"
								Case "B"
								response.write "B.60%~80%"
								Case "C"
								response.write "C.40%~60%"
								Case "D"
								response.write "D.40%以下"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>替课教师授课姓名</td>
					<td>
						<%=teachername2%>
					</td>
				</tr>
				<tr>
					<td><font color=red>停课次数</font></td>
					<td>
						<%
							Select Case stoptimes
								Case "A"
								response.write "A.0"
								Case "B"
								response.write "B.1~2"
								Case "C"
								response.write "C.3~4"
								Case "D"
								response.write "D.4以上"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>停课原因</td>
					<td>
						<%=stopcourse%>
					</td>
				</tr>
				<tr>
					<td><font color=red>停课后补课次数</font></td>
					<td>
						<%
							Select Case retimes
								Case "A"
								response.write "A.与停课次数一致"
								Case "B"
								response.write "B.少补课1次"
								Case "C"
								response.write "C.少补课2~3次"
								Case "D"
								response.write "D.少补课4次以上"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>对课程总体评价</font></td>
					<td>
						<%
							Select Case total
								Case "A"
								response.write "A.很满意"
								Case "B"
								response.write "B.比较满意"
								Case "C"
								response.write "C.尚可"
								Case "D"
								response.write "D.不太满意"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>建议：</td>
					<td>
						<textarea cols="50" rows="5" readonly><%=advice%></textarea>
					</td>
				</tr>
				

				</table>
			</td>
		</tr>
	</table>

</body>
</html>
