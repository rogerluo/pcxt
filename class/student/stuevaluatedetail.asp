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
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname")%>����ã�</td>
    <td align="right">��<a href="stuevaluate.asp">�γ��б�</a>��&nbsp;&nbsp;��<a href="stumain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="stulogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td class="titlebg" bgcolor='#cccccc'>�γ����۱�</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr bgcolor='#cccccc'>
						<td width=50%>�γ���&nbsp;&nbsp;��<%=coursename%></td>
						<td width=200>�ڿν�ʦ��<%=courseteacher%></td>
						<td>�ڿ����ͣ�<%=coursetype%></td>
					</tr>
					<tr bgcolor='#cccccc'>
						<td width=400>����<%=courseorder%>&nbsp;&nbsp;�ڿ�������<%=courseheadcount%></td>
						<td width=200>��ν�ʦ��<%=courseteacher2%></td>
						<td>ѧʱ��<%=coursetime%>&nbsp;&nbsp;ѧ�֣�<%=coursecredit%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td class="titlebg" bgcolor='#cccccc'>��������</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width=50%>�������ݣ�<font color=red>��ɫ����Ϊ��Ҫ����</font>��</td>
						<td>����ѡ���*Ϊ������д��</td>
					</tr>
					<tr>
					<td>ʹ�ý̲�</td>
					<td>
						<select name="book" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.��ʽ����̲�</option>
							<option value="B">B.�Աི��</option>
							<option value="C">C.Ӣ��Ӱӡ��</option>
							<option value="D">D.Ӣ��ԭ��̲�</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>ʹ�ÿμ�</td>
					<td>
						<select name="courseraw" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.��ʽ����μ�</option>
							<option value="B">B.�Ա�μ�</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>���ӽ̰�����</font></td>
					<td>
						<select name="rawlang" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.Ӣ��</option>
							<option value="B">B.����</option>
							<option value="C">C.��/Ӣ˫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>�ڿ�����</font></td>
					<td>
						<select name="lang" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.Ӣ��</option>
							<option value="B">B.����</option>
							<option value="C">C.��/Ӣ˫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>Ӣ���ڿεı���</font></td>
					<td>
						<select name="scale" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.Ӣ�ｲ��70%����</option>
							<option value="B">B.Ӣ�ｲ��50%~70%</option>
							<option value="C">C.Ӣ�ｲ��30%~50%</option>
							<option value="D">D.Ӣ�ｲ��30%����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>׼����֣��Ͽξ�����</td>
					<td>
						<select name="spirit" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.�ܺ�</option>
							<option value="B">B.�ȽϺ�</option>
							<option value="C">C.�п�</option>
							<option value="D">D.��̫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>��ѧ���ݳ�ʵ����ǰհ�ԣ��ṩ�㹻����Ϣ��</td>
					<td>
						<select name="info" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.�ܺ�</option>
							<option value="B">B.�ȽϺ�</option>
							<option value="C">C.�п�</option>
							<option value="D">D.��̫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>�Խ�ѧ������Ϥ���������</td>
					<td>
						<select name="teachstatus" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.�ܺ�</option>
							<option value="B">B.�ȽϺ�</option>
							<option value="C">C.�п�</option>
							<option value="D">D.��̫��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>���ν�ʦ�����ڿ�ѧʱ</font></td>
					<td>
						<select name="teachtime" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.80%����</option>
							<option value="B">B.60%~80%</option>
							<option value="C">C.40%~60%</option>
							<option value="D">D.40%����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>��ν�ʦ�ڿ�ѧʱ</font></td>
					<td>
						<select name="teachtime2" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.80%����</option>
							<option value="B">B.60%~80%</option>
							<option value="C">C.40%~60%</option>
							<option value="D">D.40%����</option>
						</select>
					</td>
				</tr>
				<tr>
					<td>��ν�ʦ�ڿ�����</td>
					<td>
						<input name="teachername2"  style="width:140"></input>
					</td>
				</tr>
				<tr>
					<td><font color=red>ͣ�δ���</font></td>
					<td>
						<select name="stoptimes" style="width:140">
							<option value="" selected="true" style="width:140"></option>
							<option value="A">A.0</option>
							<option value="B">B.1~2</option>
							<option value="C">C.3~4</option>
							<option value="D">D.4����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>ͣ��ԭ��</td>
					<td>
						<input name="stopcourse" style="width:140"></input>
					</td>
				</tr>
				<tr>
					<td><font color=red>ͣ�κ󲹿δ���</font></td>
					<td>
						<select name="retimes" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.��ͣ�δ���һ��</option>
							<option value="B">B.�ٲ���1��</option>
							<option value="C">C.�ٲ���2~3��</option>
							<option value="D">D.�ٲ���4������</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>�Կγ���������</font></td>
					<td>
						<select name="total" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.������</option>
							<option value="B">B.�Ƚ�����</option>
							<option value="C">C.�п�</option>
							<option value="D">D.��̫����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>���飺</td>
					<td>
						<textarea name="advice" cols="50" rows="5"></textarea>
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