<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	'��ȡ���Ӳ���
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))	
	stuid = Trim(request.querystring("stuid"))

	If courseid = "" Or Not IsNumeric(courseid) Or stuid = "" Or Not IsNumeric(stuid) Or coursedate = "" Then 
		conn.close()
		Set conn = Nothing 
		response.write "<script>alert('����ҳ�����');window.location.href='stuevaluate.asp';</script>"
		response.end
	End If 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о����γ�״������ϵͳ</title>
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
		response.write "<script>alert('û����Ӧ�γ���Ϣ��');window.location.href='adminlist.asp';</script>"
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
		response.write "<script>alert('û����Ӧѧ����Ϣ��');window.location.href='adminlist.asp';</script>"
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
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о����γ�״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>����ã�</td>
    <td align="right">��<a href="adminlistdetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>">������ҳ</a>��&nbsp;&nbsp;��<a href="adminmain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="adminlogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="3" cellspacing="0">
		<tr align="center" class="titlebg"><td>�γ̻�����Ϣ</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width="50%">�γ���&nbsp;&nbsp;��<%=coursename%></td>
						<td width="25%">�ڿν�ʦ��<%=courseteacher%></td>
						<td>�γ����<%=coursetype%></td>
					</tr>
					<tr>
						<td>�ڿ�������<%=courseheadcount%></td>
						<td>���&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��<%=coursetype2%></td>
						<td>ѧ��&nbsp;&nbsp;��<%=coursecredit%> ѧʱ&nbsp;&nbsp;��<%=coursetime%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td>&nbsp;&nbsp;</td></tr>
		<tr align="center" class="titlebg"><td>ѧ��������Ϣ</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width="50%">����&nbsp;&nbsp;��<%=stuname%></td>
						<td>ѧ��&nbsp;&nbsp;��<%=stuid%></td>
						
					</tr>
					<tr>
						<td width=400>רҵ&nbsp;&nbsp;��<%=stuobject%></td>
						<td width>Ժϵ&nbsp;&nbsp;��<%=stucollege%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td>&nbsp;&nbsp;</td></tr>
		<tr align="center" class="titlebg"><td>ѧ��������Ϣ</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td>�������ݣ�<font color=red>��ɫΪ��Ҫ������</font>��</td>
						<td>����ѡ���*Ϊ������д��</td>
					</tr>
					<tr>
					<td width=50%>�γ�����ƫ����</td>
					<td>
						<%
							Select Case coursecontent
								Case "A"
								response.write "A.��������"
								Case "B"
								response.write "B.����ʵ��"
								Case "C"
								response.write "C.���߼��"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>�γ�ѧϰ�Ѷ�</td>
					<td>
						<%
							Select Case coursedifficult
								Case "A"
								response.write "A.����"
								Case "B"
								response.write "B.һ��"
								Case "C"
								response.write "C.�Ƚ���"
								Case "D"
								response.write "D.����"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>�γ�����֪ʶ��</td>
					<td>
						<%
							Select Case coursepoint
								Case "A"
								response.write "A.רҵ����Դ�"
								Case "B"
								response.write "B.��ҵ����Դ�"
								Case "C"
								response.write "C.���߾���"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>�γ����ֽ���</td>
					<td>
						<%
							Select Case courselang
								Case "A"
								response.write "A.����"
								Case "B"
								response.write "B.Ӣ��"
								Case "C"
								response.write "C.��Ӣ���"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>�γ̶�����֪ʶ�����</font></td>
					<td>
						<%
							Select Case courseknowhelp
								Case "A"
								response.write "A.�ܴ�"
								Case "B"
								response.write "B.һ��"
								Case "C"
								response.write "C.����"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>�γ̶������ҹ�������</font></td>
					<td>
						<%
							Select Case coursejobhelp
								Case "A"
								response.write "A.�ܴ�"
								Case "B"
								response.write "B.һ��"
								Case "C"
								response.write "C.����"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>�γ��Ƿ����</font></td>
					<td>
						<%
							Select Case coursecontinue
								Case "A"
								response.write "A.ͬ��"
								Case "B"
								response.write "B.���뿼�췶Χ"
								Case "C"
								response.write "C.���������Ŀγ�"
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