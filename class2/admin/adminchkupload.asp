<!-- #include file="Upload.asp" -->
<!--#include file="conn.asp"-->
<!--#include file="chkstate.asp"-->
<%
'============================================
	If session("sysstate") = 1 Then 
		conn.close()
		set conn = Nothing
		response.write "<script>alert('系统尚没关闭！');window.location.href='adminmain.asp';</script>"
		response.end
	End If 
'============================================
		'获取文件名函数（不带后缀）
		Function GetFileName(ByVal strFile)
			If strFile <> "" Then
				startnum = InstrRev(strFile,"\")+1
				endnum = InstrRev(strFile,".")
				GetFileName = mid(strFile,startnum,endnum-startnum)
			Else
				GetFileName = ""
			End If
		End  Function
		'获取上传文件文件类型（即文件后缀）
		Function GetFileType(ByVal strFile)
			If strFile <> "" Then 
				GetFileType = Mid(strFile,InstrRev(strFile,".")+1)
			Else 
				GetFileType = ""
			End If 
		End Function
		'获取上传文件文件全名（包括文件后缀）
		Function GetFileFullName(ByVal strFile)
			If strFile <> "" Then
				GetFileFullName = mid(strFile,InStrRev(strFile, "\")+1)
			Else
				GetFileFullName = ""
			End If
		End  Function
'============================================
'保存form中涉及的excel文件
		formPath = "UploadFile/"
		Set upload= New DoteyUpload
		Upload.SaveTo(formPath) '将文件根据其文件名统一保存在某路径下
		If upload.ErrMsg = "" then
			response.write "<script>alert('upload sucess!');</script>"
		Else 
			response.write "<script>alert('upload failed!');</script>"	
		End If 
		Set upload=Nothing
'===========================================
		'课程计划+课程名单
		courseplan = Server.MapPath("./UploadFile") + "\courseplan.xls"
		coursemember = Server.MapPath("./UploadFile") + "\coursemember.xls"

		dim connxls,connxlsstr
	
		connxlsstr ="Provider='Microsoft.Jet.OLEDB.4.0';Extended Properties='Excel 8.0;HDR=YES;IMEX=1';Data Source="+courseplan+";"
	
		Set connxls=Server.CreateObject("ADODB.Connection") 
		Set rsxls = server.CreateObject("ADODB.Recordset")

		connxls.open connxlsstr
		strsqlxls = "select * from [Sheet1$]"
		Set rsxls = connxls.execute(strsqlxls)
		'事务处理	
		conn.BeginTrans
		
		If rsxls.eof And rsxls.bof Then
			connxls.close()
			rsxls.close()
			response.write "<script>alert('上传课程计划文件没有有效数据！');window.location.href='adminupload.asp';</script>"
			responser.end
		Else 	
			'先删除数据库中拥有相同开课时间的课程
			strsql = "delete from [courseinfo_gra] where coursedate ='"+cstr(rsxls("开课时间"))+"'"
			conn.execute strsql
								
			Do While Not rsxls.eof
			
				strsql = "insert into [courseinfo_gra] (courseid,coursename,coursetime,coursecredit,courseteacher,coursetype,coursetype2,coursedate) values ('"+CStr(rsxls("课程编号"))+"','"+CStr(rsxls("课程名称"))+"',"+(CStr(rsxls("学时")))+","+(CStr(rsxls("学分")))+",'"+CStr(rsxls("任课教师"))+"','"+CStr(rsxls("课程类别"))+"','"+CStr(rsxls("类别"))+"','"+CStr(rsxls("开课时间"))+"')"
					
				conn.execute strsql
				rsxls.movenext
			Loop
			rsxls.close()
			Set rsxls = Nothing 
			connxls.close()
			Set connxls = Nothing 
				
			If conn.Errors.Count = 0 Then
				conn.CommitTrans	
				response.write "<script>alert('upload plan success');</script>"
			Else 
				conn.RollbackTrans
				response.write "<script>alert('upload plan failed!');</script>"
			End If'判断事务是否执行成功
		End If'判断文件是否为空
			
		'选课名单
		connxlsstr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+coursemember+";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';" 
		Set connxls=Server.CreateObject("ADODB.Connection") 
		Set rsxls = server.CreateObject("ADODB.Recordset")

		connxls.open connxlsstr
		strsqlxls = "select * from [Sheet1$]"
		Set rsxls = connxls.execute(strsqlxls)

		'事务处理	
		conn.BeginTrans
		If rsxls.eof And rsxls.bof Then
			connxls.close()
			rsxls.close()
			response.write "<script>alert('上传选课名单文件没有有效数据！');window.location.href='adminupload.asp';</script>"
			responser.end
		Else 
			'先删除数据库中拥有相同开课时间的选课名单
			strsql = "delete from [coursemember_gra] where coursedate ='"+cstr(rsxls("开课时间"))+"'"
			conn.execute strsql

			Do While Not rsxls.eof
				strsql = "insert into [coursemember_gra] (stuid,stuname,stuobject,stucollege,courseid,coursedate,isevaluate) values ('"+CStr(rsxls("学号"))+"','"+CStr(rsxls("姓名"))+"','"+CStr(rsxls("专业"))+"','"+CStr(rsxls("上课系"))+"','"+CStr(rsxls("课程号"))+"','"+CStr(rsxls("开课时间"))+"',0)"

				conn.execute strsql
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
			End If'判断事务是否执行成功 
		End If'判断文件是否为空	
		
		'更新数据表
		conn.begintrans
		'同步选课名单中课程名称	
		strsql = "update [coursemember_gra] set [coursemember_gra].coursename = [courseinfo_gra].coursename from [coursemember_gra] left join [courseinfo_gra] on ([coursemember_gra].courseid = [courseinfo_gra].courseid)"
			
		conn.execute strsql
		
		'同步选课名单中课程人数
		strsql = "update [courseinfo_gra] set [courseinfo_gra].courseheadcount = (select count([coursemember_gra].stuid) from [coursemember_gra] where [coursemember_gra].courseid = [courseinfo_gra].courseid) from [courseinfo_gra]"		
			
		conn.execute strsql

		'同步选课名单中课程调查表
		strsql = "update [coursemember_gra] set [coursemember_gra].isevaluate = 1 from [coursemember_gra] join [evaluate_gra] on ([coursemember_gra].stuid = [evaluate_gra].stuid and [coursemember_gra].courseid = [evaluate_gra].courseid)"

		conn.execute strsql
		
		if conn.Errors.Count = 0 then 
			conn.committrans
			conn.close()
			set conn = nothing
			response.write "<script>history.back();</script>"
		End if	
%>