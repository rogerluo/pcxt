<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о����γ�״������ϵͳ</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{
	if(document.stuevaluate.coursecontent.value==""){
		alert("��ѡ��γ�����ƫ���ڣ�");
		document.stuevaluate.coursecontent.focus();
		return false;
	}
	if(document.stuevaluate.coursedifficult.value == ""){
		alert("��ѡ��γ�ѧϰ�Ѷȣ�");
		document.stuevaluate.coursedifficult.focus();
		return false;
	}
	if(document.stuevaluate.coursepoint.value == ""){
		alert("��ѡ��γ�����֪ʶ�㣡");
		document.stuevaluate.coursepoint.focus();
		return false;
	}
	if(document.stuevaluate.courselang.value == ""){
		alert("��ѡ��γ����ֽ��飡");
		document.stuevaluate.courselang.focus();
		return false;
	}
	if(document.stuevaluate.courseknowhelp.value == ""){
		alert("��ѡ��γ̶�������֪ʶ�������");
		document.stuevaluate.courseknowhelp.focus();
		return false;
	}
	if(document.stuevaluate.coursejobhelp.value == ""){
		alert("��ѡ��γ̶��������ҹ���������");
		document.stuevaluate.coursejobhelp.focus();
		return false;
	}
	if(document.stuevaluate.coursecontinue.value == ""){
		alert("��ѡ��γ��Ƿ������");
		document.stuevaluate.coursecontinue.focus();
		return false;
	}
}
</script>
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('����ҳ�����');history.back();</script>"
		response.end
	End If 
	
	stuid = session("stuid_gra")

	strsql = "select * from [courseinfo_gra] where courseid='"+courseid+"' and coursedate ='"+coursedate+"'"
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
	coursetype2 = CStr(rs("coursetype2"))
	
	courseheadcount = cstr(rs("courseheadcount"))
	
	conn.close
	Set conn = Nothing
%>
<body>
<form name="stuevaluate" method="post" action="stuchkevaluatedetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>" onSubmit="return CheckForm();">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о����γ�״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname_gra")%>����ã�</td>
    <td align="right">��<a href="stuevaluate.asp">�γ��б�</a>��&nbsp;&nbsp;��<a href="stumain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="stulogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td class="titlebg">�γ����۱�</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width=50%>�γ���&nbsp;&nbsp;&nbsp;&nbsp;��<%=coursename%></td>
						<td width=200>�ڿν�ʦ��<%=courseteacher%></td>
						<td>�γ����<%=coursetype%></td>
					</tr>
					<tr>
						<td width=400>�ڿ�������<%=courseheadcount%></td>
						<td width=200>���&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��<%=coursetype2%></td>
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
						<td width=50%>�������ݣ�<font color=red>��ɫ����Ϊ��Ҫ����</font>��</td>
						<td>����ѡ���*Ϊ������д��</td>
					</tr>
					<tr>
					<td>�γ�����ƫ����</td>
					<td>
						<select name="coursecontent" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.��������</option>
							<option value="B">B.����ʵ��</option>
							<option value="C">C.���߼��</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>�γ�ѧϰ�Ѷ�</td>
					<td>
						<select name="coursedifficult" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.����</option>
							<option value="B">B.һ��</option>
							<option value="C">C.�Ƚ���</option>
							<option value="D">D.����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>�γ�����֪ʶ��</td>
					<td>
						<select name="coursepoint" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.רҵ����Դ�</option>
							<option value="B">B.��ҵ����Դ�</option>
							<option value="C">C.���߾���</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td>�γ����ֽ���</td>
					<td>
						<select name="courselang" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.����</option>
							<option value="B">B.Ӣ��</option>
							<option value="C">C.��Ӣ���</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>�γ̶�������֪ʶ�����</font></td>
					<td>
						<select name="courseknowhelp" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.�ܴ�</option>
							<option value="B">B.һ��</option>
							<option value="C">C.����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>�γ̶��������ҹ�������</font></td>
					<td>
						<select name="coursejobhelp" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.�ܴ�</option>
							<option value="B">B.һ��</option>
							<option value="C">C.����</option>
						</select>*
					</td>
				</tr>
				<tr>
					<td><font color=red>�γ��Ƿ����</font></td>
					<td>
						<select name="coursecontinue" style="width:140">
							<option value="" selected="true"></option>
							<option value="A">A.ͬ��</option>
							<option value="B">B.���뿼�췶Χ</option>
							<option value="C">C.���������Ŀγ�</option>
						</select>*
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