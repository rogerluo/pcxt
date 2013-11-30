<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>电子信息工程学院研究生课程状况调查系统</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{
	if(document.stuevaluate.coursecontent.value==""){
		alert("请选择课程内容偏重于！");
		document.stuevaluate.coursecontent.focus();
		return false;
	}
	if(document.stuevaluate.coursedifficult.value == ""){
		alert("请选择课程学习难度！");
		document.stuevaluate.coursedifficult.focus();
		return false;
	}
	if(document.stuevaluate.coursepoint.value == ""){
		alert("请选择课程内容知识点！");
		document.stuevaluate.coursepoint.focus();
		return false;
	}
	if(document.stuevaluate.courselang.value == ""){
		alert("请选择课程语种建议！");
		document.stuevaluate.courselang.focus();
		return false;
	}
	if(document.stuevaluate.courseknowhelp.value == ""){
		alert("请选择课程对于自身知识面帮助！");
		document.stuevaluate.courseknowhelp.focus();
		return false;
	}
	if(document.stuevaluate.coursejobhelp.value == ""){
		alert("请选择课程对于自身找工作帮助！");
		document.stuevaluate.coursejobhelp.focus();
		return false;
	}
	if(document.stuevaluate.coursecontinue.value == ""){
		alert("请选择课程是否继续！");
		document.stuevaluate.coursecontinue.focus();
		return false;
	}
}
function ResetForm()
{
	document.stuevaluate.coursecontent.value = "";
	document.stuevaluate.coursedifficult.value = "";
	document.stuevaluate.coursepoint.value = "";
	document.stuevaluate.courselang.value = "";
	document.stuevaluate.courseknowhelp.value = "";
	document.stuevaluate.coursejobhelp.value = "";
	document.stuevaluate.coursecontinue.value = "";
}
</script>
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))

	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');history.back();</script>"
		response.end
	End If
	
	stuid = session("stuid_gra")

	strsql = "select * from [courseinfo_gra] where courseid='"+courseid+"' and coursedate='"+coursedate+"'"
	Set rs = conn.execute(strsql)

	If rs.eof Or rs.bof Then
		rs.close
		Set rs = Nothing 
		conn.close
		Set conn = Nothing 
		response.redirect "stuevaluate.asp"
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

	strsql = "select * from [evaluate_gra] where courseid='"+courseid+"' and stuid ='"+stuid+"' and coursedate ='"+coursedate+"'"
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

<form name="stuevaluate" method="post" action="stuchklistdetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>" onSubmit="return CheckForm();" onReset="return ResetForm();">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生课程状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname_gra")%>，你好！</td>
    <td align="right">【<a href="stulist.asp">查看评测</a>】&nbsp;&nbsp;【<a href="stumain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="stulogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td class="titlebg">课程评价表</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width=50%>课程名&nbsp;&nbsp;：<%=coursename%></td>
						<td width=200>授课教师：<%=courseteacher%></td>
						<td>课程类别：<%=coursetype%></td>
					</tr>
					<tr>
						<td width=400>授课人数：<%=courseheadcount%></td>
						<td width=200>类别&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;：<%=coursetype2%></td>
						<td>学时：<%=coursetime%>&nbsp;&nbsp;学分：<%=coursecredit%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td class="titlebg">评价内容</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width=50%>评价内容</td>
						<td>评价选项（带*为必须填写）</td>
					</tr>
					<tr>
					<td>课程内容偏重于</td>
					<td>
						<select name="coursecontent" style="width:140" >
							<option value=""></option>
							<option value="A" <%If coursecontent = "A" Then response.write "selected"%>>A.科研理论</option>
							<option value="B" <%If coursecontent = "B" Then response.write "selected"%>>B.具体实践</option>
							<option value="C" <%If coursecontent = "C" Then response.write "selected"%>>C.两者兼顾</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>课程学习难度</td>
					<td>
						<select name="coursedifficult" style="width:140">
							<option value=""></option>
							<option value="A" <%If coursedifficult = "A" Then response.write "selected"%>>A.容易</option>
							<option value="B" <%If coursedifficult = "B" Then response.write "selected"%>>B.一般</option>
							<option value="C" <%If coursedifficult = "C" Then response.write "selected"%>>C.比较难</option>
							<option value="D" <%If coursedifficult = "D" Then response.write "selected"%>>D.很难</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>课程内容知识点</td>
					<td>
						<select name="coursepoint" style="width:140">
							<option value=""></option>
							<option value="A" <%If coursepoint = "A" Then response.write "selected"%>>A.专业相关性大</option>
							<option value="B" <%If coursepoint = "B" Then response.write "selected"%>>B.就业相关性大</option>
							<option value="C" <%If coursepoint = "C" Then response.write "selected"%>>C.两者均否</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>课程语种建议</td>
					<td>
						<select name="courselang" style="width:140">
							<option value=""></option>
							<option value="A" <%If courselang = "A" Then response.write "selected"%>>A.中文</option>
							<option value="B" <%If courselang = "B" Then response.write "selected"%>>B.英文</option>
							<option value="C" <%If courselang = "C" Then response.write "selected"%>>C.中英结合</option>
							</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>课程对于自身知识面帮助</td>
					<td>
						<select name="courseknowhelp" style="width:140">
							<option value=""></option>
							<option value="A" <%If courseknowhelp = "A" Then response.write "selected"%>>A.很大</option>
							<option value="B" <%If courseknowhelp = "B" Then response.write "selected"%>>B.一般</option>
							<option value="C" <%If courseknowhelp = "C" Then response.write "selected"%>>C.很少</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>课程对于自身找工作帮助</td>
					<td>
						<select name="coursejobhelp" style="width:140">
							<option value=""></option>
							<option value="A" <%If coursejobhelp = "A" Then response.write "selected"%>>A.很好</option>
							<option value="B" <%If coursejobhelp = "B" Then response.write "selected"%>>B.比较好</option>
							<option value="C" <%If coursejobhelp = "C" Then response.write "selected"%>>C.尚可</option>
							<option value="D" <%If coursejobhelp = "D" Then response.write "selected"%>>D.不太好</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>课程是否继续</td>
					<td>
						<select name="coursecontinue" style="width:140">
							<option value=""></option>
							<option value="A" <%If coursecontinue = "A" Then response.write "selected"%>>A.同意</option>
							<option value="B" <%If coursecontinue = "B" Then response.write "selected"%>>B.列入考察范围</option>
							<option value="C" <%If coursecontinue = "C" Then response.write "selected"%>>C.建议更换别的课程</option>
						</select>*
					</td>
				</tr>
				<tr>
					<th colspan=2 align=center>
						<button value="submit" type="submit" name="submit">提交评价</button>
						<button type="button" value="重置" name="reset" onClick="ResetForm();">重填</button>
					</th>
				</tr>

				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>