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
		response.write "<script>alert('����ҳ�����');window.location.href='stuevaluate.asp';</script>"
		response.end
	End If 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
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
		response.write "<script>alert('û����Ӧ�γ���Ϣ��');window.location.href='adminlist.asp';</script>"
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
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>����ã�</td>
    <td align="right">��<a href="adminlistdetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>&target=0">������ҳ</a>��&nbsp;&nbsp;��<a href="adminmain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="stulogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
	<table width="800"  border="1" align="center" cellpadding="3" cellspacing="0">
		<tr align="center" class="titlebg"><td bgcolor='#cccccc'>�γ̻�����Ϣ</td></tr>
		<tr>
			<td>
				<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td width="50%">�γ���&nbsp;&nbsp;��<%=coursename%></td>
						<td width="25%">�ڿν�ʦ��<%=courseteacher%></td>
						<td>�ڿ����ͣ�<%=coursetype%></td>
					</tr>
					<tr>
						<td>�Ͽεص㣺<%=coursewhere%></td>
						<td>�ڿ�ϵ&nbsp;&nbsp;��<%=coursehost%></td>
						<td>ѧʱ&nbsp;&nbsp;&nbsp;&nbsp;��<%=coursetime%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr align="center"><td>&nbsp;&nbsp;</td></tr>
		<tr align="center" class="titlebg"><td bgcolor='#cccccc'>ѧ��������Ϣ</td></tr>
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
		<tr align="center" class="titlebg"><td bgcolor='#cccccc'>ѧ��������Ϣ</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr>
						<td>�������ݣ�<font color=red>��ɫΪ��Ҫ������</font>��</td>
						<td>����ѡ���*Ϊ������д��</td>
					</tr>
					<tr>
					<td width=50%">ʹ�ý̲�</td>
					<td>
						<%
							Select Case book
								Case "A"
								response.write "A.��ʽ����̲�"
								Case "B"
								response.write "B.�Աི��"
								Case "C"
								response.write "C.Ӣ��Ӱӡ��"
								Case "D"
								response.write "D.Ӣ��ԭ��̲�"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>ʹ�ÿμ�</td>
					<td>
						<%
							Select Case courseraw
								Case "A"
								response.write "A.��ʽ����μ�"
								Case "B"
								response.write "B.�Ա�μ�"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>���ӽ̰�����</font></td>
					<td>
						<%
							Select Case rawlang
								Case "A"
								response.write "A.Ӣ��"
								Case "B"
								response.write "B.����"
								Case "C"
								response.write "C.��/Ӣ˫��"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>�ڿ�����</font></td>
					<td>
						<%
							Select Case lang
								Case "A"
								response.write "A.Ӣ��"
								Case "B"
								response.write "B.����"
								Case "C"
								response.write "C.��/Ӣ˫��"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>Ӣ���ڿεı���</font></td>
					<td>
						<%
							Select Case scale
								Case "A"
								response.write "A.Ӣ�ｲ��70%����"
								Case "B"
								response.write "B.Ӣ�ｲ��50%~70%"
								Case "C"
								response.write "C.Ӣ�ｲ��30%~50%"
								Case "D"
								response.write "D.Ӣ�ｲ��30%����"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>׼����֣��Ͽξ�����</td>
					<td>
						<%
							Select Case spirit
								Case "A"
								response.write "A.�ܺ�"
								Case "B"
								response.write "B.�ȽϺ�"
								Case "C"
								response.write "C.�п�"
								Case "D"
								response.write "D.��̫��"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>��ѧ���ݳ�ʵ����ǰհ�ԣ��ṩ�㹻����Ϣ��</td>
					<td>
						<%
							Select Case info
								Case "A"
								response.write "A.�ܺ�"
								Case "B"
								response.write "B.�ȽϺ�"
								Case "C"
								response.write "C.�п�"
								Case "D"
								response.write "D.��̫��"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>�Խ�ѧ������Ϥ���������</td>
					<td>
						<%
							Select Case teachstatus
								Case "A"
								response.write "A.�ܺ�"
								Case "B"
								response.write "B.�ȽϺ�"
								Case "C"
								response.write "C.�п�"
								Case "D"
								response.write "D.��̫��"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>���ν�ʦ�����ڿ�ѧʱ</font></td>
					<td>
						<%
							Select Case teachtime
								Case "A"
								response.write "A.80%����"
								Case "B"
								response.write "B.60%~80%"
								Case "C"
								response.write "C.40%~60%"
								Case "D"
								response.write "D.40%����"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>��ν�ʦ�ڿ�ѧʱ</font></td>
					<td>
						<%
							Select Case teachtime2
								Case "A"
								response.write "A.80%����"
								Case "B"
								response.write "B.60%~80%"
								Case "C"
								response.write "C.40%~60%"
								Case "D"
								response.write "D.40%����"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>��ν�ʦ�ڿ�����</td>
					<td>
						<%=teachername2%>
					</td>
				</tr>
				<tr>
					<td><font color=red>ͣ�δ���</font></td>
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
								response.write "D.4����"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>ͣ��ԭ��</td>
					<td>
						<%=stopcourse%>
					</td>
				</tr>
				<tr>
					<td><font color=red>ͣ�κ󲹿δ���</font></td>
					<td>
						<%
							Select Case retimes
								Case "A"
								response.write "A.��ͣ�δ���һ��"
								Case "B"
								response.write "B.�ٲ���1��"
								Case "C"
								response.write "C.�ٲ���2~3��"
								Case "D"
								response.write "D.�ٲ���4������"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td><font color=red>�Կγ���������</font></td>
					<td>
						<%
							Select Case total
								Case "A"
								response.write "A.������"
								Case "B"
								response.write "B.�Ƚ�����"
								Case "C"
								response.write "C.�п�"
								Case "D"
								response.write "D.��̫����"
							End Select 
						%>
					</td>
				</tr>
				<tr>
					<td>���飺</td>
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
