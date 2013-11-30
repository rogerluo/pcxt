<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>电子信息工程学院研究生教学状况调查系统</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{
	if(document.stuevaluate.book.value==""){
		alert("请选择使用教材！");
		document.stuevaluate.book.focus();
		return false;
	}
	if(document.stuevaluate.courseraw.value == ""){
		alert("请选择使用课件！");
		document.stuevaluate.courseraw.focus();
		return false;
	}
	if(document.stuevaluate.rawlang.value == ""){
		alert("请选择电子教案语种！");
		document.stuevaluate.rawlang.focus();
		return false;
	}
	if(document.stuevaluate.lang.value == ""){
		alert("请选择授课语种！");
		document.stuevaluate.lang.focus();
		return false;
	}
	if(document.stuevaluate.scale.value == ""){
		alert("请选择英语授课比例！");
		document.stuevaluate.scale.focus();
		return false;
	}
	if(document.stuevaluate.spirit.value == ""){
		alert("请选择上课精神饱满情况！");
		document.stuevaluate.spirit.focus();
		return false;
	}
	if(document.stuevaluate.info.value == ""){
		alert("请选择授课信息量情况！");
		document.stuevaluate.info.focus();
		return false;
	}
	if(document.stuevaluate.teachstatus.value == ""){
		alert("请选择讲述清楚情况！");
		document.stuevaluate.teachstatus.focus();
		return false;
	}
	if(document.stuevaluate.teachtime.value == ""){
		alert("请选择开课教师授课学时情况！");
		document.stuevaluate.teachtime.focus();
		return false;
	}
	if(document.stuevaluate.total.value == ""){
		alert("请选择课题总体评价！");
		document.stuevaluate.total.focus();
		return false;
	}
	if(document.stuevaluate.total.value == ""){
		alert("请选择课题总体评价！");
		document.stuevaluate.total.focus();
		return false;
	}
	/*
	if(document.stuevaluate.teachtime2.value!="")
	{
		if(document.stuevaluate.teachername2.value == "")
		{
			alert("请填写替课教师姓名！");
			document.stuevaluate.teachername2.focus();
			return false;
		}
	}
	*/
	if(document.stuevaluate.teachername2.value!="")
	{
		if(document.stuevaluate.teachtime2.value == "")
		{
			alert("请选择替课教师授课学时！");
			document.stuevaluate.teachtime2.focus();
			return false;
		}
	}
	if (document.stuevaluate.stoptimes.value == "")
	{
			alert("请选择停课次数！");
			document.stuevaluate.stoptimes.focus();
			return false;			
	}
	
	if (document.stuevaluate.retimes.value == "")
	{
			alert("请选择补课次数！");
			document.stuevaluate.retimes.focus();
			return false;			
	}
	if (document.stuevaluate.stoptimes.value < document.stuevaluate.retimes.value)
	{
			alert("请正确选择停课与补课次数！");
			document.stuevaluate.retimes.focus();
			return false;			
	}	
	/*
	if (document.stuevaluate.stoptimes.value != ""||document.stuevaluate.retimes.value!=""||document.stuevaluate.stopcourse.value!="")
	{
	
		if(document.stuevaluate.stoptimes.value == "")
		{
			alert("请选择听课次数！");
			document.stuevaluate.stoptimes.focus();
			return false;
		}
		if(document.stuevaluate.retimes.value == "")
		{
			alert("请选择补课次数！");
			document.stuevaluate.retimes.focus();
			return false;
		}
		if (document.stuevaluate.stoptimes.value < document.stuevaluate.retimes.value == "")
		{
			alert("请选择正确")	
		}
	}
	*/
	if(document.stuevaluate.teachername2.value.length>10)
	{
		alert("替课教师名字过长（10字以内）！");
		document.stuevaluate.teachername2.value.focus();
		return false;
	}
	if(document.stuevaluate.advice.value.length>100)
	{
		alert("建议内容过长（100字以内）！");
		document.stuevaluate.advice.value.focus();
		return false;
	}
	
}
</script>
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	courseorder = Trim(request.querystring("courseorder"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" or courseorder = "" or Not IsNumeric(courseorder) Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');history.back();</script>"
		response.end
	End If 	
	stuid = session("stuid")

	strsql = "select * from [courseinfo] where courseid='"+courseid+"' and coursedate='"+coursedate+"' and courseorder="+courseorder
	Set rs = conn.execute(strsql)
	If rs.eof Or rs.bof Then
		rs.close
		Set rs = Nothing 
		conn.close
		Set conn = Nothing 
		response.redirect "stuevaluate.asp"
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

	strsql = "select * from [evaluate] where courseid='"+courseid+"' and stuid ='"+stuid+"' and coursedate ='"+coursedate+"' and courseorder="+courseorder
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

<form name="stuevaluate" method="post" action="stuchklistdetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>" onSubmit="return CheckForm();">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生教学状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname")%>，你好！</td>
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
						<td>授课类型：<%=coursetype%></td>
					</tr>
					<tr>
						<td width=400>课序：<%=courseorder%>&nbsp;&nbsp;授课人数：<%=courseheadcount%></td>
						<td width=200>替课教师：<%=courseteacher2%></td>
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
					<td>使用教材</td>
					<td>
						<select name="book" style="width:140" >
							<option value=""></option>
							<option value="A" <%If book = "A" Then response.write "selected"%>>A.正式出版教材</option>
							<option value="B" <%If book = "B" Then response.write "selected"%>>B.自编讲义</option>
							<option value="C" <%If book = "C" Then response.write "selected"%>>C.英文影印版</option>
							<option value="D" <%If book = "D" Then response.write "selected"%>>D.英文原版教材</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>使用课件</td>
					<td>
						<select name="courseraw" style="width:140">
							<option value=""></option>
							<option value="A" <%If courseraw = "A" Then response.write "selected"%>>A.正式出版课件</option>
							<option value="B" <%If courseraw = "B" Then response.write "selected"%>>B.自编课件</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>电子教案语种</td>
					<td>
						<select name="rawlang" style="width:140">
							<option value=""></option>
							<option value="A" <%If rawlang = "A" Then response.write "selected"%>>A.英文</option>
							<option value="B" <%If rawlang = "B" Then response.write "selected"%>>B.中文</option>
							<option value="C" <%If rawlang = "C" Then response.write "selected"%>>C.中/英双语</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>授课语种</td>
					<td>
						<select name="lang" style="width:140">
							<option value=""></option>
							<option value="A" <%If lang = "A" Then response.write "selected"%>>A.英文</option>
							<option value="B" <%If lang = "B" Then response.write "selected"%>>B.中文</option>
							<option value="C" <%If lang = "C" Then response.write "selected"%>>C.中/英双语</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>英语授课的比例</td>
					<td>
						<select name="scale" style="width:140">
							<option value=""></option>
							<option value="A" <%If scale = "A" Then response.write "selected"%>>A.英语讲述70%以上</option>
							<option value="B" <%If scale = "B" Then response.write "selected"%>>B.英语讲述50%~70%</option>
							<option value="C" <%If scale = "C" Then response.write "selected"%>>C.英语讲述30%~50%</option>
							<option value="D" <%If scale = "D" Then response.write "selected"%>>D.英语讲述30%以下</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>准备充分，上课精神饱满</td>
					<td>
						<select name="spirit" style="width:140">
							<option value=""></option>
							<option value="A" <%If spirit = "A" Then response.write "selected"%>>A.很好</option>
							<option value="B" <%If spirit = "B" Then response.write "selected"%>>B.比较好</option>
							<option value="C" <%If spirit = "C" Then response.write "selected"%>>C.尚可</option>
							<option value="D" <%If spirit = "D" Then response.write "selected"%>>D.不太好</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>教学内容充实具有前瞻性，提供足够的信息量</td>
					<td>
						<select name="info" style="width:140">
							<option value=""></option>
							<option value="A" <%If info = "A" Then response.write "selected"%>>A.很好</option>
							<option value="B" <%If info = "B" Then response.write "selected"%>>B.比较好</option>
							<option value="C" <%If info = "C" Then response.write "selected"%>>C.尚可</option>
							<option value="D" <%If info = "D" Then response.write "selected"%>>D.不太好</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>对教学内容熟悉，讲述清楚</td>
					<td>
						<select name="teachstatus" style="width:140">
							<option value=""></option>
							<option value="A" <%If teachstatus = "A" Then response.write "selected"%>>A.很好</option>
							<option value="B" <%If teachstatus = "B" Then response.write "selected"%>>B.比较好</option>
							<option value="C" <%If teachstatus = "C" Then response.write "selected"%>>C.尚可</option>
							<option value="D" <%If teachstatus = "D" Then response.write "selected"%>>D.不太好</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>开课教师本人授课学时</td>
					<td>
						<select name="teachtime" style="width:140">
							<option value=""></option>
							<option value="A" <%If teachtime = "A" Then response.write "selected"%>>A.80%以上</option>
							<option value="B" <%If teachtime = "B" Then response.write "selected"%>>B.60%~80%</option>
							<option value="C" <%If teachtime = "C" Then response.write "selected"%>>C.40%~60%</option>
							<option value="D" <%If teachtime = "D" Then response.write "selected"%>>D.40%以下</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>替课教师授课学时</td>
					<td>
						<select name="teachtime2" style="width:140">
							<option value=""></option>
							<option value="A" <%If teachtime2 = "A" Then response.write "selected"%>>A.80%以上</option>
							<option value="B" <%If teachtime2 = "B" Then response.write "selected"%>>B.60%~80%</option>
							<option value="C" <%If teachtime2 = "C" Then response.write "selected"%>>C.40%~60%</option>
							<option value="D" <%If teachtime2 = "D" Then response.write "selected"%>>D.40%以下</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>替课教师授课姓名</td>
					<td>
						<input name="teachername2"  style="width:140" value="<%=teachername2%>"></input>
					</td>
				</tr>
				<tr>
					<td>停课次数</td>
					<td>
						<select name="stoptimes" style="width:140">
							<option value=""></option>
							<option value="A" <%If stoptimes = "A" Then response.write "selected"%>>A.0</option>
							<option value="B" <%If stoptimes = "B" Then response.write "selected"%>>B.1~2</option>
							<option value="C" <%If stoptimes = "C" Then response.write "selected"%>>C.3~4</option>
							<option value="D" <%If stoptimes = "D" Then response.write "selected"%>>D.4以上</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>停课原因</td>
					<td>
						<input name="stopcourse" style="width:140" value="<%=stopcourse%>"></input>
					</td>
				</tr>
				<tr>
					<td>停课后补课次数</td>
					<td>
						<select name="retimes" style="width:140">
							<option value=""></option>
							<option value="A" <%If retimes = "A" Then response.write "selected"%>>A.与停课次数一致</option>
							<option value="B" <%If retimes = "B" Then response.write "selected"%>>B.少补课1次</option>
							<option value="C" <%If retimes = "C" Then response.write "selected"%>>C.少补课2~3次</option>
							<option value="D" <%If retimes = "D" Then response.write "selected"%>>D.少补课4次以上</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>对课程总体评价</td>
					<td>
						<select name="total" style="width:140">
							<option value=""></option>
							<option value="A" <%If total = "A" Then response.write "selected"%>>A.很满意</option>
							<option value="B" <%If total = "B" Then response.write "selected"%>>B.比较满意</option>
							<option value="C" <%If total = "C" Then response.write "selected"%>>C.尚可</option>
							<option value="D" <%If total = "D" Then response.write "selected"%>>D.不太满意</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>建议：</td>
					<td>
						<textarea name="advice" cols="50" rows="5"><%=advice%></textarea>
					</td>
				</tr>
				<tr>
					<th colspan=2 align=center>
						<button value="submit" type="submit" name="submit">提交评价</button>
						<button type="reset" value="重置">重填</button>
					</th>
				</tr>

				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>