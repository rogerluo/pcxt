<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{
	if(document.stuevaluate.book.value==""){
		alert("��ѡ��ʹ�ý̲ģ�");
		document.stuevaluate.book.focus();
		return false;
	}
	if(document.stuevaluate.courseraw.value == ""){
		alert("��ѡ��ʹ�ÿμ���");
		document.stuevaluate.courseraw.focus();
		return false;
	}
	if(document.stuevaluate.rawlang.value == ""){
		alert("��ѡ����ӽ̰����֣�");
		document.stuevaluate.rawlang.focus();
		return false;
	}
	if(document.stuevaluate.lang.value == ""){
		alert("��ѡ���ڿ����֣�");
		document.stuevaluate.lang.focus();
		return false;
	}
	if(document.stuevaluate.scale.value == ""){
		alert("��ѡ��Ӣ���ڿα�����");
		document.stuevaluate.scale.focus();
		return false;
	}
	if(document.stuevaluate.spirit.value == ""){
		alert("��ѡ���Ͽξ����������");
		document.stuevaluate.spirit.focus();
		return false;
	}
	if(document.stuevaluate.info.value == ""){
		alert("��ѡ���ڿ���Ϣ�������");
		document.stuevaluate.info.focus();
		return false;
	}
	if(document.stuevaluate.teachstatus.value == ""){
		alert("��ѡ������������");
		document.stuevaluate.teachstatus.focus();
		return false;
	}
	if(document.stuevaluate.teachtime.value == ""){
		alert("��ѡ�񿪿ν�ʦ�ڿ�ѧʱ�����");
		document.stuevaluate.teachtime.focus();
		return false;
	}
	if(document.stuevaluate.total.value == ""){
		alert("��ѡ������������ۣ�");
		document.stuevaluate.total.focus();
		return false;
	}
	if(document.stuevaluate.total.value == ""){
		alert("��ѡ������������ۣ�");
		document.stuevaluate.total.focus();
		return false;
	}
	/*
	if(document.stuevaluate.teachtime2.value!="")
	{
		if(document.stuevaluate.teachername2.value == "")
		{
			alert("����д��ν�ʦ������");
			document.stuevaluate.teachername2.focus();
			return false;
		}
	}
	*/
	if(document.stuevaluate.teachername2.value!="")
	{
		if(document.stuevaluate.teachtime2.value == "")
		{
			alert("��ѡ����ν�ʦ�ڿ�ѧʱ��");
			document.stuevaluate.teachtime2.focus();
			return false;
		}
	}
	if (document.stuevaluate.stoptimes.value == "")
	{
			alert("��ѡ��ͣ�δ�����");
			document.stuevaluate.stoptimes.focus();
			return false;			
	}
	
	if (document.stuevaluate.retimes.value == "")
	{
			alert("��ѡ�񲹿δ�����");
			document.stuevaluate.retimes.focus();
			return false;			
	}
	if (document.stuevaluate.stoptimes.value < document.stuevaluate.retimes.value)
	{
			alert("����ȷѡ��ͣ���벹�δ�����");
			document.stuevaluate.retimes.focus();
			return false;			
	}	
	/*
	if (document.stuevaluate.stoptimes.value != ""||document.stuevaluate.retimes.value!=""||document.stuevaluate.stopcourse.value!="")
	{
	
		if(document.stuevaluate.stoptimes.value == "")
		{
			alert("��ѡ�����δ�����");
			document.stuevaluate.stoptimes.focus();
			return false;
		}
		if(document.stuevaluate.retimes.value == "")
		{
			alert("��ѡ�񲹿δ�����");
			document.stuevaluate.retimes.focus();
			return false;
		}
		if (document.stuevaluate.stoptimes.value < document.stuevaluate.retimes.value == "")
		{
			alert("��ѡ����ȷ")	
		}
	}
	*/
	if(document.stuevaluate.teachername2.value.length>10)
	{
		alert("��ν�ʦ���ֹ�����10�����ڣ���");
		document.stuevaluate.teachername2.value.focus();
		return false;
	}
	if(document.stuevaluate.advice.value.length>100)
	{
		alert("�������ݹ�����100�����ڣ���");
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
		response.write "<script>alert('����ҳ�����');history.back();</script>"
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
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname")%>����ã�</td>
    <td align="right">��<a href="stulist.asp">�鿴����</a>��&nbsp;&nbsp;��<a href="stumain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="stulogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td class="titlebg">�γ����۱�</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width=50%>�γ���&nbsp;&nbsp;��<%=coursename%></td>
						<td width=200>�ڿν�ʦ��<%=courseteacher%></td>
						<td>�ڿ����ͣ�<%=coursetype%></td>
					</tr>
					<tr>
						<td width=400>����<%=courseorder%>&nbsp;&nbsp;�ڿ�������<%=courseheadcount%></td>
						<td width=200>��ν�ʦ��<%=courseteacher2%></td>
						<td>ѧʱ��<%=coursetime%>&nbsp;&nbsp;ѧ�֣�<%=coursecredit%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td class="titlebg">��������</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width=50%>��������</td>
						<td>����ѡ���*Ϊ������д��</td>
					</tr>
					<tr>
					<td>ʹ�ý̲�</td>
					<td>
						<select name="book" style="width:140" >
							<option value=""></option>
							<option value="A" <%If book = "A" Then response.write "selected"%>>A.��ʽ����̲�</option>
							<option value="B" <%If book = "B" Then response.write "selected"%>>B.�Աི��</option>
							<option value="C" <%If book = "C" Then response.write "selected"%>>C.Ӣ��Ӱӡ��</option>
							<option value="D" <%If book = "D" Then response.write "selected"%>>D.Ӣ��ԭ��̲�</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>ʹ�ÿμ�</td>
					<td>
						<select name="courseraw" style="width:140">
							<option value=""></option>
							<option value="A" <%If courseraw = "A" Then response.write "selected"%>>A.��ʽ����μ�</option>
							<option value="B" <%If courseraw = "B" Then response.write "selected"%>>B.�Ա�μ�</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>���ӽ̰�����</td>
					<td>
						<select name="rawlang" style="width:140">
							<option value=""></option>
							<option value="A" <%If rawlang = "A" Then response.write "selected"%>>A.Ӣ��</option>
							<option value="B" <%If rawlang = "B" Then response.write "selected"%>>B.����</option>
							<option value="C" <%If rawlang = "C" Then response.write "selected"%>>C.��/Ӣ˫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>�ڿ�����</td>
					<td>
						<select name="lang" style="width:140">
							<option value=""></option>
							<option value="A" <%If lang = "A" Then response.write "selected"%>>A.Ӣ��</option>
							<option value="B" <%If lang = "B" Then response.write "selected"%>>B.����</option>
							<option value="C" <%If lang = "C" Then response.write "selected"%>>C.��/Ӣ˫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>Ӣ���ڿεı���</td>
					<td>
						<select name="scale" style="width:140">
							<option value=""></option>
							<option value="A" <%If scale = "A" Then response.write "selected"%>>A.Ӣ�ｲ��70%����</option>
							<option value="B" <%If scale = "B" Then response.write "selected"%>>B.Ӣ�ｲ��50%~70%</option>
							<option value="C" <%If scale = "C" Then response.write "selected"%>>C.Ӣ�ｲ��30%~50%</option>
							<option value="D" <%If scale = "D" Then response.write "selected"%>>D.Ӣ�ｲ��30%����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>׼����֣��Ͽξ�����</td>
					<td>
						<select name="spirit" style="width:140">
							<option value=""></option>
							<option value="A" <%If spirit = "A" Then response.write "selected"%>>A.�ܺ�</option>
							<option value="B" <%If spirit = "B" Then response.write "selected"%>>B.�ȽϺ�</option>
							<option value="C" <%If spirit = "C" Then response.write "selected"%>>C.�п�</option>
							<option value="D" <%If spirit = "D" Then response.write "selected"%>>D.��̫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>��ѧ���ݳ�ʵ����ǰհ�ԣ��ṩ�㹻����Ϣ��</td>
					<td>
						<select name="info" style="width:140">
							<option value=""></option>
							<option value="A" <%If info = "A" Then response.write "selected"%>>A.�ܺ�</option>
							<option value="B" <%If info = "B" Then response.write "selected"%>>B.�ȽϺ�</option>
							<option value="C" <%If info = "C" Then response.write "selected"%>>C.�п�</option>
							<option value="D" <%If info = "D" Then response.write "selected"%>>D.��̫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>�Խ�ѧ������Ϥ���������</td>
					<td>
						<select name="teachstatus" style="width:140">
							<option value=""></option>
							<option value="A" <%If teachstatus = "A" Then response.write "selected"%>>A.�ܺ�</option>
							<option value="B" <%If teachstatus = "B" Then response.write "selected"%>>B.�ȽϺ�</option>
							<option value="C" <%If teachstatus = "C" Then response.write "selected"%>>C.�п�</option>
							<option value="D" <%If teachstatus = "D" Then response.write "selected"%>>D.��̫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>���ν�ʦ�����ڿ�ѧʱ</td>
					<td>
						<select name="teachtime" style="width:140">
							<option value=""></option>
							<option value="A" <%If teachtime = "A" Then response.write "selected"%>>A.80%����</option>
							<option value="B" <%If teachtime = "B" Then response.write "selected"%>>B.60%~80%</option>
							<option value="C" <%If teachtime = "C" Then response.write "selected"%>>C.40%~60%</option>
							<option value="D" <%If teachtime = "D" Then response.write "selected"%>>D.40%����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>��ν�ʦ�ڿ�ѧʱ</td>
					<td>
						<select name="teachtime2" style="width:140">
							<option value=""></option>
							<option value="A" <%If teachtime2 = "A" Then response.write "selected"%>>A.80%����</option>
							<option value="B" <%If teachtime2 = "B" Then response.write "selected"%>>B.60%~80%</option>
							<option value="C" <%If teachtime2 = "C" Then response.write "selected"%>>C.40%~60%</option>
							<option value="D" <%If teachtime2 = "D" Then response.write "selected"%>>D.40%����</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>��ν�ʦ�ڿ�����</td>
					<td>
						<input name="teachername2"  style="width:140" value="<%=teachername2%>"></input>
					</td>
				</tr>
				<tr>
					<td>ͣ�δ���</td>
					<td>
						<select name="stoptimes" style="width:140">
							<option value=""></option>
							<option value="A" <%If stoptimes = "A" Then response.write "selected"%>>A.0</option>
							<option value="B" <%If stoptimes = "B" Then response.write "selected"%>>B.1~2</option>
							<option value="C" <%If stoptimes = "C" Then response.write "selected"%>>C.3~4</option>
							<option value="D" <%If stoptimes = "D" Then response.write "selected"%>>D.4����</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>ͣ��ԭ��</td>
					<td>
						<input name="stopcourse" style="width:140" value="<%=stopcourse%>"></input>
					</td>
				</tr>
				<tr>
					<td>ͣ�κ󲹿δ���</td>
					<td>
						<select name="retimes" style="width:140">
							<option value=""></option>
							<option value="A" <%If retimes = "A" Then response.write "selected"%>>A.��ͣ�δ���һ��</option>
							<option value="B" <%If retimes = "B" Then response.write "selected"%>>B.�ٲ���1��</option>
							<option value="C" <%If retimes = "C" Then response.write "selected"%>>C.�ٲ���2~3��</option>
							<option value="D" <%If retimes = "D" Then response.write "selected"%>>D.�ٲ���4������</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>�Կγ���������</td>
					<td>
						<select name="total" style="width:140">
							<option value=""></option>
							<option value="A" <%If total = "A" Then response.write "selected"%>>A.������</option>
							<option value="B" <%If total = "B" Then response.write "selected"%>>B.�Ƚ�����</option>
							<option value="C" <%If total = "C" Then response.write "selected"%>>C.�п�</option>
							<option value="D" <%If total = "D" Then response.write "selected"%>>D.��̫����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>���飺</td>
					<td>
						<textarea name="advice" cols="50" rows="5"><%=advice%></textarea>
					</td>
				</tr>
				<tr>
					<th colspan=2 align=center>
						<button value="submit" type="submit" name="submit">�ύ����</button>
						<button type="reset" value="����">����</button>
					</th>
				</tr>

				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>