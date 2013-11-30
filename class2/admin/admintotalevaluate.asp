<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	'获取链接参数
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');window.location.href='adminlist.asp';</script>"
		response.end
	End If 

	strsql = "select * from [courseinfo_gra] where courseid ='"+courseid+"' and coursedate = '"+coursedate+"'"
	Set rs = conn.execute(strsql)

	coursename = CStr(rs("coursename"))
	coursetime = CStr(rs("coursetime"))
	coursecredit = cstr(rs("coursecredit"))
	courseteacher = cstr(rs("courseteacher"))
	coursetype = cstr(rs("coursetype"))
	coursetype2 = cstr(rs("coursetype2"))

	courseheadcount = CStr(rs("courseheadcount"))
	totalnum = courseheadcount

	rs.close() 

	strsql = "select count(stuid) as evaluated from [coursemember_gra] where courseid='"+courseid+"' and isevaluate = 1 and coursedate='"+coursedate+"'"
	Set rs = conn.execute(strsql)

	evaluated = rs("evaluated")
	
	rs.close
	'获取一共有几个人支持
	coursestudycount = 0
	coursejobcount = 0

	strsql = "select [evaluate2_gra].stuid,[evaluate2_gra].stuname,[evaluate2_gra].coursestudy,[evaluate2_gra].coursejob from [coursemember_gra] join [evaluate2_gra] on ([coursemember_gra].stuid = [evaluate2_gra].stuid and [coursemember_gra].courseid = '"+courseid+"' and [coursemember_gra].coursedate = '"+coursedate+"')"
	Set rs = conn.execute(strsql)

	If rs.eof Or rs.bof Then 
	Else
		Do While Not rs.eof
			coursestudystr = rs("coursestudy")
			coursejobstr = rs("coursejob")
			coursestudyarray = Split(coursestudystr,",")
			coursejobarray = Split(coursejobstr,",")
			For i = 0 To UBound(coursestudyarray)
				If trim(courseid) = Trim(coursestudyarray(i)) Then
					coursestudycount = coursestudycount + 1
					Exit For 
				End If 
			Next
			For i = 0 To UBound(coursejobarray)
				If Trim(coursejobarray(i)) = Trim(courseid) Then
					coursejobcount = coursejobcount + 1
					Exit For 
				End If 
			Next
			rs.movenext
		Loop 
	End If 

	rs.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META   HTTP-EQUIV="Pragma"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Cache-Control"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Expires"   CONTENT="0">    
<title>电子信息工程学院研究生课程状况调查系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
<%
  Response.Buffer   =   True    
  Response.ExpiresAbsolute   =   Now()   -   1    
  Response.Expires   =   0    
  Response.CacheControl   =   "no-cache"    
  Response.AddHeader   "Pragma",   "No-Cache"  
%>
</head>
<%
	response.write "<script language=javascript>"
	response.write "function Pop(){"
	response.write "window.open ('adminpop.asp?target=0&courseid="+courseid+"&coursedate="+coursedate+"', 'helloworld', 'height=400, width=400, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');}</script>"
	response.write "<script language=javascript>"
	response.write "function Pop2(){"
	response.write "window.open ('adminpop.asp?target=1&courseid="+courseid+"&coursedate="+coursedate+"', 'helloworld', 'height=400, width=400, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no');}</script>"
	
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
<table width="800" border="1" align="center" cellpadding="5" cellspacing="0" class="tbbg">
 <tr> 
      <td align="center" class="titlebg">课程基本信息</td>
 </tr>
 <tr>
	<td>
		<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
			<tr>
				<td width=50%>课程名&nbsp;&nbsp;：<%=coursename%></td>
				<td width=25%>授课教师：<%=courseteacher%></td>
				<td>课程类别：<%=coursetype%></td>
			</tr>
			<tr>
				<td>授课人数：<%=courseheadcount%></td>
				<td width=25%>类别&nbsp;&nbsp;：<%=coursetype2%></td>
				<td>学时：<%=coursetime%>&nbsp;&nbsp;学分：<%=coursecredit%></td>
			</tr>
		</table>
	</td>
 </tr>
  <tr> 
      <td align="center" class="titlebg">评价统计信息</td>
 </tr>
  <tr> 
      <td> 
	  		<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
				<tr>
					<td width="33%">
					选修课程人数：共&nbsp;&nbsp;<%=CStr(totalnum)%>&nbsp;&nbsp;人
					</td>
					<td width="33%">
					已评学生人数：共&nbsp;&nbsp;<%=CStr(evaluated)%>&nbsp;&nbsp;人【<a href="adminunevaluate.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&target=0">查看</a>】
					</td>
					<td>
					尚没评价人数：共&nbsp;&nbsp;<%=CStr(totalnum-evaluated)%>&nbsp;&nbsp;人【<a href="adminunevaluate.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&target=1">查看</a>】
					</td>
				</tr>
			</table>
	  </td>
 </tr>
   <tr>
	<td>
		<table width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
			<tr>
				<td width=400>
					对自身深造有帮助的课程：共&nbsp;&nbsp;<%=coursestudycount%>&nbsp;&nbsp;人&nbsp;&nbsp;
					<input type=button name=button1 value="查看" onClick="return Pop();" style="width:100">
				</td>
				<td>
					对自身工作有帮助的课程：共&nbsp;&nbsp;<%=coursejobcount%>&nbsp;&nbsp;人&nbsp;&nbsp;<input type=button name=button2 value="查看" onClick="return Pop2();" style="width:100">
				</td>
			</tr>
		</table>
	</td>
   </tr>
   <tr> 
      <td> 
<%
	Dim itemarray(8,5)
	strsql = "select * from [evaluate_gra] where courseid='"+courseid+"' and coursedate='"+coursedate+"'"
	Set rs = conn.execute(strsql)

	If rs.eof And rs.bof Then 
	Else 
		Do While Not rs.eof
			For i = 1 To 7
				Select Case rs(i+4)
					Case "A"
					itemarray(i,1) = itemarray(i,1) + 1
					Case "B"
					itemarray(i,2) = itemarray(i,2) + 1
					Case "C"
					itemarray(i,3) = itemarray(i,3) + 1
					Case "D"
					itemarray(i,4) = itemarray(i,4) + 1
				End Select
			Next	
			rs.movenext
		Loop
	End If 
	rs.close
	For i = 1 To 7 
		If itemarray(i,1) = "" Then 
			itemarray(i,1) = 0
		End If
		If itemarray(i,2) = "" Then 
			itemarray(i,2) = 0
		End If
		If itemarray(i,3) = "" Then 
			itemarray(i,3) = 0
		End If
		If itemarray(i,4) = "" Then 
			itemarray(i,4) = 0
		End If
	Next 
	
	'========避免错误
	If evaluated = 0 Then 
		evaluated = 1
	End If 
%>
	  		<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
				<tr align=center>
					<td width=400>评价表项（<font color=red>红色为主要考查项</font>）</td>
					<td width=12%>A</td>
					<td width=12%>B</td>
					<td width=12%>C</td>
					<td>D</td>
					

				</tr>
				<tr align=left>
					<td>课程内容偏重于</td>
					<td><%=CStr(itemarray(1,1))+"("+CStr(FormatNumber(itemarray(1,1)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(1,2))+"("+CStr(FormatNumber(itemarray(1,2)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(1,3))+"("+CStr(FormatNumber(itemarray(1,3)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(1,4))+"("+CStr(FormatNumber(itemarray(1,4)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					
				</tr>
				<tr>
					<td>课程学习难度</td>
					<td><%=CStr(itemarray(2,1))+"("+CStr(FormatNumber(itemarray(2,1)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(2,2))+"("+CStr(FormatNumber(itemarray(2,2)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(2,3))+"("+CStr(FormatNumber(itemarray(2,3)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(2,4))+"("+CStr(FormatNumber(itemarray(2,4)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					
				</tr>
				<tr>
					<td>课程内容知识点</td>
					<td><%=CStr(itemarray(3,1))+"("+CStr(FormatNumber(itemarray(3,1)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(3,2))+"("+CStr(FormatNumber(itemarray(3,2)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(3,3))+"("+CStr(FormatNumber(itemarray(3,3)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(3,4))+"("+CStr(FormatNumber(itemarray(3,4)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					
				</tr>
				<tr>
					<td>课程语种建议</td>
					<td><%=CStr(itemarray(4,1))+"("+CStr(FormatNumber(itemarray(4,1)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(4,2))+"("+CStr(FormatNumber(itemarray(4,2)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(4,3))+"("+CStr(FormatNumber(itemarray(4,3)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(4,4))+"("+CStr(FormatNumber(itemarray(4,4)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					
				</tr>
				<tr>
					<td><font color=red>课程对于自身知识面帮助</font></td>
					<td><%=CStr(itemarray(5,1))+"("+CStr(FormatNumber(itemarray(5,1)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(5,2))+"("+CStr(FormatNumber(itemarray(5,2)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(5,3))+"("+CStr(FormatNumber(itemarray(5,3)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(5,4))+"("+CStr(FormatNumber(itemarray(5,4)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					
				</tr>
				<tr>
					<td><font color=red>课程对于自身找工作帮助</font></td>
					<td><%=CStr(itemarray(6,1))+"("+CStr(FormatNumber(itemarray(6,1)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(6,2))+"("+CStr(FormatNumber(itemarray(6,2)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(6,3))+"("+CStr(FormatNumber(itemarray(6,3)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(6,4))+"("+CStr(FormatNumber(itemarray(6,4)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					
				</tr>
				<tr>
					<td><font color=red>课程是否继续</font></td>
					<td><%=CStr(itemarray(7,1))+"("+CStr(FormatNumber(itemarray(7,1)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(7,2))+"("+CStr(FormatNumber(itemarray(7,2)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(7,3))+"("+CStr(FormatNumber(itemarray(7,3)/evaluated*100,2,-1,-1,0))+"%)"%></td>
					<td><%=CStr(itemarray(7,4))+"("+CStr(FormatNumber(itemarray(7,4)/evaluated*100,2,-1,-1,0))+"%)"%></td>
				</tr>
			</table>
	  </td>
 </tr>
</table>
</body>
</html>
