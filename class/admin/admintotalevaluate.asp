<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	courseorder = Trim(request.querystring("courseorder"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" or courseorder = "" or Not IsNumeric(courseorder) Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');window.location.href='adminlist.asp';</script>"
		response.end
	End If 

	strsql = "select * from [courseinfo] where courseid ='"+courseid+"' and coursedate = '"+coursedate+"' and courseorder = "+courseorder
	Set rs = conn.execute(strsql)
	If rs.eof Or rs.bof Then
		rs.close
		response.write "<script>alert('请求页面错误！');window.location.href='adminlist.asp';</script>"
		response.end
	Else 
		coursename = rs("coursename")
		coursetime = CStr(rs("coursetime"))
		courseteacher = cstr(rs("courseteacher"))
		coursetype = cstr(rs("coursetype"))
		courseheadcount = cstr(rs("courseheadcount"))
		coursecredit = cstr(rs("coursecredit"))
		courseteacher2 = cstr(rs("courseteacher2"))
		totalnum = CInt(courseheadcount)
	End If 
	rs.close() 

	strsql = "select count(stuid) as evaluated from [coursemember] where courseid='"+courseid+"' and isevaluate =1 and coursedate='"+coursedate+"' and courseorder =" + courseorder
	Set rs = conn.execute(strsql)
	evaluated = rs("evaluated")
	rs.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META   HTTP-EQUIV="Pragma"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Cache-Control"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Expires"   CONTENT="0">    
<title>电子信息工程学院研究生教学状况调查系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
<%
  Response.Buffer   =   True    
  Response.ExpiresAbsolute   =   Now()   -   1    
  Response.Expires   =   0    
  Response.CacheControl   =   "no-cache"    
  Response.AddHeader   "Pragma",   "No-Cache"  

%>
</head>
<script language='javascript'>
function Output()
{
	window.location.href='adminoutput.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>';
}
</script>
<body>

<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生教学状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>，你好！</td>
    <td align="right">【<a href="adminlistdetail.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>">返回上页</a>】&nbsp;&nbsp;【<a href="adminmain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="adminlogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
<table width="800" border="1" align="center" cellpadding="5" cellspacing="0" class="tbbg"">
 <tr> 
      <td align="center" class="titlebg" bgcolor='#cccccc'>课程基本信息</td>
 </tr>
 <tr>
	<td>
		<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
			<tr>
				<td width=400>课程名&nbsp;&nbsp;：<%=coursename%></td>
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
  <tr> 
      <td align="center" class="titlebg" bgcolor='#cccccc'>评价统计信息</td>
 </tr>
  <tr> 
      <td> 
	  		<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
				<tr>
					<td width="33%">

					选修课程人数：共&nbsp;&nbsp;<%=CStr(totalnum)%>&nbsp;&nbsp;人
					</td>
					<td width="33%">
					已评学生人数：共&nbsp;&nbsp;<%=CStr(evaluated)%>&nbsp;&nbsp;人【<a href="adminunevaluate.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>&target=0">查看</a>】
					</td>
					<td>
					尚没评价人数：共&nbsp;&nbsp;<%=CStr(totalnum-evaluated)%>&nbsp;&nbsp;人【<a href="adminunevaluate.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>&target=1">查看</a>】
					</td>
				</tr>
			</table>
	  </td>
 </tr>
   <tr> 
      <td> 
<%
	Dim itemarray(17,5)
	strsql = "select * from [evaluate] where courseid='"+courseid+"' and coursedate='"+coursedate+"' and courseorder =" + courseorder
	Set rs = conn.execute(strsql)
	If rs.eof And rs.bof Then 
	
	Else 
		Do While Not rs.eof
			For i = 1 To 16
				'response.write rs(i+3)+"<P>"
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
	For i = 1 To 16 
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
					<td width=50%>评价表项（<font color=red>红色为主要考查项</font>）</td>
					<td width=10%>A</td>
					<td width=10%>B</td>
					<td width=10%>C</td>
					<td width=10%>D</td>
					

				</tr>
				<tr align=left>
					<td>使用教材</td>
					<td><%
					optionA=CStr(itemarray(1,1))+"("+CStr(FormatNumber(itemarray(1,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA	
					%></td>
					<td><%
					optionB=CStr(itemarray(1,2))+"("+CStr(FormatNumber(itemarray(1,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(1,3))+"("+CStr(FormatNumber(itemarray(1,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(1,4))+"("+CStr(FormatNumber(itemarray(1,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'清除totalevaluate表内的数据
	strsql = "delete from totalevaluate"
	conn.execute(strsql)
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',1,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>					
				</tr>
				<tr>
					<td>使用课件</td>
					<td><%
					optionA=CStr(itemarray(2,1))+"("+CStr(FormatNumber(itemarray(2,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(2,2))+"("+CStr(FormatNumber(itemarray(2,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(2,3))+"("+CStr(FormatNumber(itemarray(2,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(2,4))+"("+CStr(FormatNumber(itemarray(2,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',2,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>
				</tr>
				<tr>
					<td><font color=red>电子教案语种</font></td>
					<td><%
					optionA=CStr(itemarray(3,1))+"("+CStr(FormatNumber(itemarray(3,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(3,2))+"("+CStr(FormatNumber(itemarray(3,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(3,3))+"("+CStr(FormatNumber(itemarray(3,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(3,4))+"("+CStr(FormatNumber(itemarray(3,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+"',3,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>
					
				</tr>
				<tr>
					<td><font color=red>授课语种</font></td>
					<td><%
					optionA=CStr(itemarray(4,1))+"("+CStr(FormatNumber(itemarray(4,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(4,2))+"("+CStr(FormatNumber(itemarray(4,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(4,3))+"("+CStr(FormatNumber(itemarray(4,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(4,4))+"("+CStr(FormatNumber(itemarray(4,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',4,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>					
				</tr>
				<tr>
					<td><font color=red>英语授课的比例</font></td>
					<td><%
					optionA=CStr(itemarray(5,1))+"("+CStr(FormatNumber(itemarray(5,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(5,2))+"("+CStr(FormatNumber(itemarray(5,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(5,3))+"("+CStr(FormatNumber(itemarray(5,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(5,4))+"("+CStr(FormatNumber(itemarray(5,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',5,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>							
				</tr>
				<tr>
					<td>准备充分，上课精神饱满</td>
					<td><%
					optionA=CStr(itemarray(6,1))+"("+CStr(FormatNumber(itemarray(6,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(6,2))+"("+CStr(FormatNumber(itemarray(6,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(6,3))+"("+CStr(FormatNumber(itemarray(6,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(6,4))+"("+CStr(FormatNumber(itemarray(6,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',6,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>								
				</tr>
				<tr>
					<td>提供足够的信息量</td>
					<td><%
					optionA=CStr(itemarray(7,1))+"("+CStr(FormatNumber(itemarray(7,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(7,2))+"("+CStr(FormatNumber(itemarray(7,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(7,3))+"("+CStr(FormatNumber(itemarray(7,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(7,4))+"("+CStr(FormatNumber(itemarray(7,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',7,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>									
				</tr>
				<tr>
					<td>对教学内容熟悉，讲述清楚</td>
					<td><%
					optionA=CStr(itemarray(8,1))+"("+CStr(FormatNumber(itemarray(8,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(8,2))+"("+CStr(FormatNumber(itemarray(8,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(8,3))+"("+CStr(FormatNumber(itemarray(8,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(8,4))+"("+CStr(FormatNumber(itemarray(8,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',8,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>								
				</tr>
				<tr>
					<td><font color=red>开课教师本人授课学时</font></td>
					<td><%
					optionA=CStr(itemarray(9,1))+"("+CStr(FormatNumber(itemarray(9,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(9,2))+"("+CStr(FormatNumber(itemarray(9,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(9,3))+"("+CStr(FormatNumber(itemarray(9,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(9,4))+"("+CStr(FormatNumber(itemarray(9,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',9,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>							
				</tr>
				<tr>
					<td><font color=red>停课次数</font></td>
					<td><%
					optionA=CStr(itemarray(12,1))+"("+CStr(FormatNumber(itemarray(12,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(12,2))+"("+CStr(FormatNumber(itemarray(12,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(12,3))+"("+CStr(FormatNumber(itemarray(12,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(12,4))+"("+CStr(FormatNumber(itemarray(12,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',10,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>							
				</tr>
				<tr>
					<td><font color=red>停课后补课次数</font></td>
					<td><%
					optionA=CStr(itemarray(14,1))+"("+CStr(FormatNumber(itemarray(14,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(14,2))+"("+CStr(FormatNumber(itemarray(14,2)/evaluated*100,2,-1,-1,0))+"%)"				
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(14,3))+"("+CStr(FormatNumber(itemarray(14,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(14,4))+"("+CStr(FormatNumber(itemarray(14,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',11,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>								
				</tr>				
				<tr>
					<td><font color=red>对课程总体评价</font></td>
					<td><%
					optionA=CStr(itemarray(15,1))+"("+CStr(FormatNumber(itemarray(15,1)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionA
					%></td>
					<td><%
					optionB=CStr(itemarray(15,2))+"("+CStr(FormatNumber(itemarray(15,2)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionB
					%></td>
					<td><%
					optionC=CStr(itemarray(15,3))+"("+CStr(FormatNumber(itemarray(15,3)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionC
					%></td>
					<td><%
					optionD=CStr(itemarray(15,4))+"("+CStr(FormatNumber(itemarray(15,4)/evaluated*100,2,-1,-1,0))+"%)"
					response.write optionD
					%></td>
<%
	'更新totalevaluate表内的数据
	strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+courseid+"',"+courseorder+",'"+coursedate+"',12,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
	conn.execute(strsql)
%>						
				</tr>
			</table>
	  </td>
 </tr>
 <tr>
	<td align="center" class="titlebg" bgcolor='#cccccc'>
		<input type='button' name='output' value='导出表格' style='width:100' onclick="return Output();"></input>
	</td>
 </tr>
</table>
</body>
</html>