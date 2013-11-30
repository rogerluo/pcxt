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

	strsql = "select * from [courseinfo] where courseid='"+courseid+"' and coursedate ='"+coursedate+"' and courseorder = "+courseorder
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
	conn.close
	Set conn = Nothing
%>
<body>
<form name="stuevaluate" method="post" action="stuchkevaluatedetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>" onSubmit="return CheckForm();">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生教学状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname")%>，你好！</td>
    <td align="right">【<a href="stuevaluate.asp">课程列表</a>】&nbsp;&nbsp;【<a href="stumain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="stulogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td class="titlebg" bgcolor='#cccccc'>课程评价表</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr bgcolor='#cccccc'>
						<td width=50%>课程名&nbsp;&nbsp;：<%=coursename%></td>
						<td width=200>授课教师：<%=courseteacher%></td>
						<td>授课类型：<%=coursetype%></td>
					</tr>
					<tr bgcolor='#cccccc'>
						<td width=400>课序：<%=courseorder%>&nbsp;&nbsp;授课人数：<%=courseheadcount%></td>
						<td width=200>替课教师：<%=courseteacher2%></td>
						<td>学时：<%=coursetime%>&nbsp;&nbsp;学分：<%=coursecredit%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td class="titlebg" bgcolor='#cccccc'>评价内容</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width=50%>评价内容（<font color=red>红色内容为主要考查</font>）</td>
						<td>评价选项（带*为必须填写）</td>
					</tr>
					<tr>
					<td>使用教材</td>
					<td>
						<select name="book" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.正式出版教材</option>
							<option value="B">B.自编讲义</option>
							<option value="C">C.英文影印版</option>
							<option value="D">D.英文原版教材</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>使用课件</td>
					<td>
						<select name="courseraw" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.正式出版课件</option>
							<option value="B">B.自编课件</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>电子教案语种</font></td>
					<td>
						<select name="rawlang" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.英文</option>
							<option value="B">B.中文</option>
							<option value="C">C.中/英双语</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>授课语种</font></td>
					<td>
						<select name="lang" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.英文</option>
							<option value="B">B.中文</option>
							<option value="C">C.中/英双语</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>英语授课的比例</font></td>
					<td>
						<select name="scale" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.英语讲述70%以上</option>
							<option value="B">B.英语讲述50%~70%</option>
							<option value="C">C.英语讲述30%~50%</option>
							<option value="D">D.英语讲述30%以下</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>准备充分，上课精神饱满</td>
					<td>
						<select name="spirit" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.很好</option>
							<option value="B">B.比较好</option>
							<option value="C">C.尚可</option>
							<option value="D">D.不太好</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>教学内容充实具有前瞻性，提供足够的信息量</td>
					<td>
						<select name="info" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.很好</option>
							<option value="B">B.比较好</option>
							<option value="C">C.尚可</option>
							<option value="D">D.不太好</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>对教学内容熟悉，讲述清楚</td>
					<td>
						<select name="teachstatus" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.很好</option>
							<option value="B">B.比较好</option>
							<option value="C">C.尚可</option>
							<option value="D">D.不太好</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>开课教师本人授课学时</font></td>
					<td>
						<select name="teachtime" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.80%以上</option>
							<option value="B">B.60%~80%</option>
							<option value="C">C.40%~60%</option>
							<option value="D">D.40%以下</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>替课教师授课学时</font></td>
					<td>
						<select name="teachtime2" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.80%以上</option>
							<option value="B">B.60%~80%</option>
							<option value="C">C.40%~60%</option>
							<option value="D">D.40%以下</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>替课教师授课姓名</td>
					<td>
						<input name="teachername2"  style="width:140"></input>
					</td>
				</tr>
				<tr>
					<td><font color=red>停课次数</font></td>
					<td>
						<select name="stoptimes" style="width:140">
							<option value="" selected="true" style="width:140"></option>
							<option value="A">A.0</option>
							<option value="B">B.1~2</option>
							<option value="C">C.3~4</option>
							<option value="D">D.4以上</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>停课原因</td>
					<td>
						<input name="stopcourse" style="width:140"></input>
					</td>
				</tr>
				<tr>
					<td><font color=red>停课后补课次数</font></td>
					<td>
						<select name="retimes" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.与停课次数一致</option>
							<option value="B">B.少补课1次</option>
							<option value="C">C.少补课2~3次</option>
							<option value="D">D.少补课4次以上</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>对课程总体评价</font></td>
					<td>
						<select name="total" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.很满意</option>
							<option value="B">B.比较满意</option>
							<option value="C">C.尚可</option>
							<option value="D">D.不太满意</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>建议：</td>
					<td>
						<textarea name="advice" cols="50" rows="5"></textarea>
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