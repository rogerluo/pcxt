<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="connxls.asp"-->
<%
	If session("sysstate") = 1 Then 
		conn.close()
		set conn = Nothing
		response.write "<script>alert('ϵͳ��û�رգ�');window.location.href='adminmain.asp';</script>"
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

	
	act = "upload"
	If act = "upload" Then		
	'�ж��Ƿ��д�ֵ'
		'��ȡ����'
		courseid = Trim(request.Form("courseid"))
		coursename = Trim(request.Form("coursename"))
		coursetime = Trim(request.Form("coursetime"))
		courseteacher = Trim(request.Form("courseteacher"))
		coursetype = Trim(request.Form("coursetype"))
		coursehost = Trim(request.Form("coursehost"))
		coursewhere = Trim(request.Form("coursewhere"))
		coursefile = Trim(request.Form("coursefile"))
		'�ж��Ƿ�Ϊ��Чֵ'
		If courseid = "" Or coursename = "" Or coursetime = "" Or courseteacher = "" Or coursetype = "" Or coursehost = "" Or coursewhere = "" Or coursefile = "" Then
			conn.close()
			set conn = Nothing
			response.write "<script>alert('�ϴ��ļ�ʧ�ܣ�');window.location.href='adminupload.asp';</script>"
			response.end
		End If
		filetype = GetFileType(coursefile)
		If filetype = "xls" Or filetype = "XLS" Or filetype = "Xls" Then
		Else 
			conn.close()
			set conn = Nothing
			response.write "<script>alert('�ϴ��ļ����Ͳ�����Ҫ��');window.location.href='adminupload.asp';</script>"			
			response.end
		End If 
		'�����ļ�'		
		Set objStream = Server.CreateObject("ADODB.Stream")
		objStream.Type = 1 
		objStream.Open
		objStream.LoadFromFile coursefile
		savedfile = Server.MapPath("../excels/")+"\"+(GetFileName(coursefile))+"_by_"+session("adminname")+"_"+CStr(Date())+"."+(GetFileType(coursefile))
		objStream.SaveToFile savedfile,2
		'objStream.SaveToFile Server.MapPath("../excels/")+"\"+(GetFileFullName(coursefile)),2
		objStream.Close
		'�������ݱ�courseinfo'
		'���Ȳ�ѯ�Ƿ�������Ӧ�Ŀγ̺�'
		strsql = "select courseid from [courseinfo] where courseid = '"+courseid +"'"
		Set rs = conn.execute(strsql)
		If rs.eof And rs.bof Then			
		Else 
			'������ͬ��ɾ��֮ǰ���ڱ�[courseinfo],[coursemember]�������Ϣ'
			Set conn2 = server.CreateObject("ADODB.Connection")
			conn2.open connstr
			conn2.BeginTrans
			'strsql2 = "delete [courseinfo],[coursemember] from [courseinfo],[coursemember] where [coursemember].courseid = '" + courseid + "' and [courseinfo].courseid ='" + courseid + "'"
			strsql2 = "delete from [courseinfo] where [courseinfo].courseid = '" + courseid +"'"
			conn2.execute(strsql2)		
			If conn2.Errors.Count = 0 Then
				conn2.CommitTrans
				response.write("Delete [courseinfo] success<P>")
			Else 
				conn2.RollbackTrans
				response.write("Delete [courseinfo] failed<p>")
			End If 

			conn2.close()
			Set conn2 = Nothing

			Set conn3 = server.CreateObject("ADODB.Connection")
			conn3.open connstr
			conn3.BeginTrans
			strsql3 = "delete from [coursemember] where [coursemember].courseid = '" + courseid +"'"
			conn3.execute(strsql3)
			If conn3.Errors.Count = 0 Then
				conn3.CommitTrans
				response.write("Delete [coursemember] success<P>")
			Else 
				conn3.RollbackTrans
				response.write("Delete [coursemember] failed<p>")
			End If 

			conn3.close()
			Set conn3 = Nothing
		End If
			'�������ݱ�[courseinfo]'
			Set conn4 = server.CreateObject("ADODB.Connection")
			conn4.open connstr
			conn4.BeginTrans
			strsql4 = "insert [courseinfo] (courseid,coursename,coursetime,courseteacher,coursetype,coursehost,coursewhere,coursefile) values ('"+courseid+"','"+coursename+"',"+coursetime+",'"+courseteacher+"','"+coursetype+"','"+coursehost+"','"+coursewhere+"','"+savedfile+"')"
			'strsql3 = "insert [courseinfo] (courseid,coursename,coursetime) values ('"+courseid+"','"+coursename+"',"+(coursetime)+")"
			conn4.execute(strsql4)
			If conn4.Errors.Count = 0 Then
				conn4.CommitTrans
				response.write("Insert [courseinfo] success<P>")
			Else 
				conn4.RollbackTrans
				response.write("Insert [courseinfo] failed<p>")
			End If 
			conn4.close()
			Set conn4 = Nothing
			
			'�������ݱ�[coursemember]'
			'�򿪸ձ����excel�ļ�'
			connxlsstr = connxlsstr + "DBQ="+savedfile
			connxls.open connxlsstr
			strsqlxls = "select * from [Sheet1$]"
			Set rsxls = connxls.execute(strsqlxls)
			'��[coursemember]����'
			Set conn5 = server.CreateObject("ADODB.Connection")
			conn5.open connstr
			conn5.BeginTrans

			If rsxls.eof And rsxls.bof Then
				connxls.close()
				rsxls.close()
				conn5.close()
				rs.close()
				conn.close()
				Set conn = Nothing 
				response.write "<script>alert('�ϴ��ļ�û����Ч���ݣ�');window.location.href='adminupload.asp';</script>"
				responser.end
			Else 
				Do While Not rsxls.eof
					strsql5 = "insert [coursemember] (stuid,stuname,stuobject,stucollege,courseid,coursename,isevaluate) values ('"+CStr(rsxls("ѧ��"))+"','"+CStr(rsxls("����"))+"','"+CStr(rsxls("רҵ"))+"','"+CStr(rsxls("�Ͽ�ϵ"))+"','"+courseid+"','"+coursename+"',0)"
					'strsql3 = "insert [coursemember] (courseid) values('"+CStr(courseid)+"')"
					conn5.execute strsql5
					rsxls.movenext
				Loop
				rsxls.close()
				Set rsxls = Nothing 
				connxls.close()
				Set connxls = Nothing 
				If conn5.Errors.Count = 0 Then
					conn5.CommitTrans
					response.write("Insert [coursemember] success<P>")
				Else 
					conn5.RollbackTrans
					response.write("Insert [coursemember] failed<p>")
				End If 
				conn5.close()
				Set conn5 = Nothing 
			End If 
		rs.close()
		Set rs = Nothing 
		conn.close()
		Set conn = Nothing
		'�������ݿ�ɹ�'
		response.write "<script>alert('�ϴ��ļ��ɹ���');window.location.href='adminupload.asp';</script>"
	End If 
%>