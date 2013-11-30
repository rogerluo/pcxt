<!-- #include file="Upload.asp" -->
<!--#include file="conn.asp"-->
<!--#include file="chkstate.asp"-->
<%
	If session("sysstate") = 1 Then 
		conn.close()
		set conn = Nothing
		response.write "<script>alert('系统尚没关闭！');window.location.href='adminmain.asp';</script>"
		response.end
	End If 
'==============
		Function GetFileName(ByVal strFile)
			If strFile <> "" Then
				startnum = InstrRev(strFile,"\")+1
				endnum = InstrRev(strFile,".")
				GetFileName = mid(strFile,startnum,endnum-startnum)
			Else
				GetFileName = ""
			End If
		End  Function
		Function GetFileType(ByVal strFile)
			If strFile <> "" Then 
				GetFileType = Mid(strFile,InstrRev(strFile,".")+1)
			Else 
				GetFileType = ""
			End If 
		End Function
		Function GetFileFullName(ByVal strFile)
			If strFile <> "" Then
				GetFileFullName = mid(strFile,InStrRev(strFile, "\")+1)
			Else
				GetFileFullName = ""
			End If
		End  Function
'=================
	'coursefile = Trim(request.Form("coursefile"))
	'coursemember = Trim(request.Form("coursemember"))
	'判断是否为有效值'
	'If coursefile = "" or coursemember = "" Then
	'	conn.close()
	'	set conn = Nothing
	'	response.write "<script>alert('上传文件失败！');window.location.href='adminupload.asp';</script>"
	'	response.end
	'End If
	'filetype = GetFileType(coursefile)
	'If filetype = "xls" Or filetype = "XLS" Or filetype = "Xls" Then
	'Else 
	'		conn.close()
	'		set conn = Nothing
	'		response.write "<script>alert('上传课程计划文件类型不符合要求！');window.location.href='adminupload.asp';</script>"			
	'		response.end
	'End If 
	'filetype = GetFileType(coursemember)
	'If filetype = "xls" Or filetype = "XLS" Or filetype = "Xls" Then
	'Else 
	'		conn.close()
	'		set conn = Nothing
	'		response.write "<script>alert('上传选课名单文件类型不符合要求！');window.location.href='adminupload.asp';</script>"			
	'		response.end
	'End If 
		'复制文件'课程计划
		'Server.ScriptTimeout = 900
		

		formPath = "UploadFiles/"
		Set upload= New DoteyUpload
		Upload.SaveTo(formPath) '将文件根据其文件名统一保存在某路径下
		If upload.ErrMsg = "" then
			response.write "<script>alert('upload sucess!');</script>"
		Else 
			response.write "<script>alert('upload failed!');</script>"	
		End If 
		Set upload=Nothing

		'课程计划 课程名单
		coursefile = Server.MapPath("./UploadFiles") + "\courseplan.xls"
		coursemember = Server.MapPath("./UploadFiles") + "\coursemember.xls"

		'Set objStream = Server.CreateObject("ADODB.Stream")
		'objStream.Type = 1 
		'objStream.Open
		'objStream.LoadFromFile coursefile
		'savedfile = Server.MapPath("../excels/")+"\"+(GetFileName(coursefile))+"_by_"+session("adminname")+"_"+CStr(Date())+"."+(GetFileType(coursefile))	
		'objStream.SaveToFile savedfile,2
		'objStream.Close
		'复制文件'选课名单
		'Set objStream = Server.CreateObject("ADODB.Stream")
		'objStream.Type = 1 
		'objStream.Open
		'objStream.LoadFromFile coursemember
		'savedfile2 = Server.MapPath("../excels/")+"\"+(GetFileName(coursemember))+"_by_"+session("adminname")+"_"+CStr(Date())+"."+(GetFileType(coursemember))
		'objStream.SaveToFile savedfile2,2
		'objStream.Close

	dim connxls,connxlsstr
	'savedfile = "C:\Inetpub\class\excels\my_by_razorluo_2010-3-25.xls"
	'connxlsstr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+coursefile+";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';"
	connxlsstr ="Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Excel 8.0;HDR=YES;IMEX=1';Data Source="+coursefile+";"
	'connxlsstr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+coursefile+";Extended Properties=Excel 8.0;"
	'response.write (coursefile+"<p>"+connxlsstr+"<p>")
	Set connxls=Server.CreateObject("ADODB.Connection") 
	Set rsxls = server.CreateObject("ADODB.Recordset")
'strsqlxls = "select * from [Sheet1$]"
'Set rsxls = connxls.execute(strsqlxls)
'		If rsxls.eof And rsxls.bof Then
'		Else
'			Do While Not rsxls.eof
'				response.write CStr(rsxls(1)) + "   " + CStr(rsxls(2)) +"    "+CStr(rsxls("课程号："))+ "<p>"	
'				rsxls.movenext
'			Loop
'		End If
'			rsxls.close()
'			Set rsxls = Nothing
'			connxls.close()
'			Set connxls = Nothing
			'connxlsstr = connxlsstr + "DBQ="+savedfile
			connxls.open connxlsstr
			strsqlxls = "select * from [Sheet1$]"
			Set rsxls = connxls.execute(strsqlxls)
			'打开[coursemember]连接'
			'Set conn3 = server.CreateObject("ADODB.Connection")
			'conn3.open connstr
			'conn3.BeginTrans
			totalclass = 0
			totalmember = 0
			conn.BeginTrans
			If rsxls.eof And rsxls.bof Then
				connxls.close()
				rsxls.close()
				'conn3.close()
				'conn.close()
				'Set conn = Nothing 
				response.write "<script>alert('上传课程计划文件没有有效数据！');window.location.href='adminupload.asp';</script>"
				responser.end
				Else 
			
			'===============
			'根据开课时间来删除[courseinfo]中已有的课程，从而保证唯一性
			strsql = "delete from [courseinfo] where coursedate ='"+cstr(rsxls("开课时间"))+"'"
			conn.execute strsql
			
			'===============				
				
				
				Do While Not rsxls.eof
					'response.write CSTr(rsxls(0)) +"<p>"
					'strsql = "insert into [courseinfo] (courseid) values('"+CStr(rsxls(0))+"')"
					
					If IsNull(rsxls(0)) Then
						response.write "null"
						EXIT DO	
					End If
					teacher2 = rsxls("教师2")
					If IsNull(teacher2) Then 
						teacher2 = ""
					End If 
					'response.write (CStr(rsxls("课程编号"))+" "+CStr(rsxls("课程名称"))+" "+CStr(rsxls("教师1"))+" "+CStr(teacher2)+" "+CStr(rsxls("课程类型"))+" "+CStr(rsxls("开课时间"))+"<p>")
					strsql = "insert into [courseinfo] (courseid,coursename,coursetime,coursecredit,courseteacher,courseteacher2,coursetype,coursefile,coursedate, courseorder, courseheadcount) values ('"+CStr(rsxls("课程编号"))+"','"+CStr(rsxls("课程名称"))+"',"+CStr(rsxls("学时"))+","+CStr(rsxls("学分"))++",'"+CStr(rsxls("教师1"))+"','"+CStr(teacher2)+"','"+CStr(rsxls("课程类型"))+"','"+savedfile+"','"+CStr(rsxls("开课时间"))+"', "+cstr(rsxls("课序"))+",0)"
					'strsql3 = "insert [coursemember] (stuid,stuname,stuobject,stucollege,courseid,coursename,isevaluate) values ('"+CStr(rsxls("学号"))+"','"+CStr(rsxls("姓名"))+"','"+CStr(rsxls("专业"))+"','"+CStr(rsxls("上课系"))+"','"+courseid+"','"+coursename+"',0)"
					'Set rs3 = conn3.execute(strsql3)
					conn.execute strsql
					totalcourse = totalcourse + 1
					rsxls.movenext
				Loop
				rsxls.close()
				Set rsxls = Nothing 
				connxls.close()
				Set connxls = Nothing 
				If conn.Errors.Count = 0 Then
					conn.CommitTrans
					'response.write("Update [courseinfo] success<P>")
					response.write "<script>alert('upload plan success');</script>"
				Else 
					conn.RollbackTrans
					response.write "<script>alert('upload plan failed!');</script>"
				End If 
				'conn.close()
				'Set conn = Nothing 
				'response.write "<script>alert('后面！');history.back();</script>"' 没有执行
			End If 
			
			'选课名单
			connxlsstr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+coursemember+";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';" 
			Set connxls=Server.CreateObject("ADODB.Connection") 
			Set rsxls = server.CreateObject("ADODB.Recordset")

			connxls.open connxlsstr
			strsqlxls = "select * from [Sheet1$]"
			Set rsxls = connxls.execute(strsqlxls)

			conn.BeginTrans
			If rsxls.eof And rsxls.bof Then
				connxls.close()
				rsxls.close()
				'conn3.close()
				'conn.close()
				'Set conn = Nothing 
				response.write "<script>alert('上传选课名单文件没有有效数据！');window.location.href='adminupload.asp';</script>"
				responser.end
				Else 
			
			'===============
			'根据开课时间来删除[coursemember]中已有的课程，从而保证唯一性
			strsql = "delete from [coursemember] where coursedate ='"+cstr(rsxls("开课时间"))+"'"
			conn.execute strsql
			
			'===============								
				Do While Not rsxls.eof
					strsql = "insert into [coursemember] (stuid,stuname,stuobject,stucollege,courseid,coursedate,courseorder,isevaluate) values ('"+CStr(rsxls("学号"))+"','"+CStr(rsxls("姓名"))+"','"+CStr(rsxls("专业"))+"','"+CStr(rsxls("上课系"))+"','"+CStr(rsxls("课程号"))+"','"+CStr(rsxls("开课时间"))+"',"+cstr(rsxls("课序"))+",0)"
					conn.execute strsql
					totalmember = totalmember + 1
					rsxls.movenext
				Loop
				rsxls.close()
				Set rsxls = Nothing 
				connxls.close()
				Set connxls = Nothing 
				If conn.Errors.Count = 0 Then
					conn.CommitTrans
					response.write "<script>alert('upload member success！');</script>"
				Else 
					conn.RollbackTrans
					response.write "<script>alert('upload member failed！');</script>"
				End If 
				'conn.close()
				'Set conn = Nothing 
				
			End If 	
			response.write "<script>alert('共读取了 "+cstr(totalcourse)+" 门课, "+cstr(totalmember)+" 条上课记录, 请确认数字是否符合结果');</script>"
			'更新数据表
			conn.begintrans
			strsql = "update [coursemember] set [coursemember].coursename = [courseinfo].coursename from [coursemember] left join [courseinfo] on ([coursemember].courseid = [courseinfo].courseid and [coursemember].coursedate = [courseinfo].coursedate) and [coursemember].courseorder = [courseinfo].courseorder"
			conn.execute strsql
			strsql = "update [courseinfo] set [courseinfo].courseheadcount = (select count([coursemember].stuid) from [coursemember] where [coursemember].courseid = [courseinfo].courseid and [coursemember].coursedate = [courseinfo].coursedate and [coursemember].courseorder = [courseinfo].courseorder) from [courseinfo]"		
			conn.execute strsql
			strsql = "update [coursemember] set [coursemember].isevaluate = 1 from [coursemember] join [evaluate] on ([coursemember].stuid = [evaluate].stuid and [coursemember].courseid = [evaluate].courseid)"
			conn.execute strsql
			if conn.Errors.Count = 0 then 
				conn.committrans
				conn.close()
				set conn = nothing
				response.write "<script>history.back();</script>"
			end if
			
%>
