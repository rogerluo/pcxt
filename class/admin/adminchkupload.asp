<!-- #include file="Upload.asp" -->
<!--#include file="conn.asp"-->
<!--#include file="chkstate.asp"-->
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
	'coursefile = Trim(request.Form("coursefile"))
	'coursemember = Trim(request.Form("coursemember"))
	'�ж��Ƿ�Ϊ��Чֵ'
	'If coursefile = "" or coursemember = "" Then
	'	conn.close()
	'	set conn = Nothing
	'	response.write "<script>alert('�ϴ��ļ�ʧ�ܣ�');window.location.href='adminupload.asp';</script>"
	'	response.end
	'End If
	'filetype = GetFileType(coursefile)
	'If filetype = "xls" Or filetype = "XLS" Or filetype = "Xls" Then
	'Else 
	'		conn.close()
	'		set conn = Nothing
	'		response.write "<script>alert('�ϴ��γ̼ƻ��ļ����Ͳ�����Ҫ��');window.location.href='adminupload.asp';</script>"			
	'		response.end
	'End If 
	'filetype = GetFileType(coursemember)
	'If filetype = "xls" Or filetype = "XLS" Or filetype = "Xls" Then
	'Else 
	'		conn.close()
	'		set conn = Nothing
	'		response.write "<script>alert('�ϴ�ѡ�������ļ����Ͳ�����Ҫ��');window.location.href='adminupload.asp';</script>"			
	'		response.end
	'End If 
		'�����ļ�'�γ̼ƻ�
		'Server.ScriptTimeout = 900
		

		formPath = "UploadFiles/"
		Set upload= New DoteyUpload
		Upload.SaveTo(formPath) '���ļ��������ļ���ͳһ������ĳ·����
		If upload.ErrMsg = "" then
			response.write "<script>alert('upload sucess!');</script>"
		Else 
			response.write "<script>alert('upload failed!');</script>"	
		End If 
		Set upload=Nothing

		'�γ̼ƻ� �γ�����
		coursefile = Server.MapPath("./UploadFiles") + "\courseplan.xls"
		coursemember = Server.MapPath("./UploadFiles") + "\coursemember.xls"

		'Set objStream = Server.CreateObject("ADODB.Stream")
		'objStream.Type = 1 
		'objStream.Open
		'objStream.LoadFromFile coursefile
		'savedfile = Server.MapPath("../excels/")+"\"+(GetFileName(coursefile))+"_by_"+session("adminname")+"_"+CStr(Date())+"."+(GetFileType(coursefile))	
		'objStream.SaveToFile savedfile,2
		'objStream.Close
		'�����ļ�'ѡ������
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
'				response.write CStr(rsxls(1)) + "   " + CStr(rsxls(2)) +"    "+CStr(rsxls("�γ̺ţ�"))+ "<p>"	
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
			'��[coursemember]����'
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
				response.write "<script>alert('�ϴ��γ̼ƻ��ļ�û����Ч���ݣ�');window.location.href='adminupload.asp';</script>"
				responser.end
				Else 
			
			'===============
			'���ݿ���ʱ����ɾ��[courseinfo]�����еĿγ̣��Ӷ���֤Ψһ��
			strsql = "delete from [courseinfo] where coursedate ='"+cstr(rsxls("����ʱ��"))+"'"
			conn.execute strsql
			
			'===============				
				
				
				Do While Not rsxls.eof
					'response.write CSTr(rsxls(0)) +"<p>"
					'strsql = "insert into [courseinfo] (courseid) values('"+CStr(rsxls(0))+"')"
					
					If IsNull(rsxls(0)) Then
						response.write "null"
						EXIT DO	
					End If
					teacher2 = rsxls("��ʦ2")
					If IsNull(teacher2) Then 
						teacher2 = ""
					End If 
					'response.write (CStr(rsxls("�γ̱��"))+" "+CStr(rsxls("�γ�����"))+" "+CStr(rsxls("��ʦ1"))+" "+CStr(teacher2)+" "+CStr(rsxls("�γ�����"))+" "+CStr(rsxls("����ʱ��"))+"<p>")
					strsql = "insert into [courseinfo] (courseid,coursename,coursetime,coursecredit,courseteacher,courseteacher2,coursetype,coursefile,coursedate, courseorder, courseheadcount) values ('"+CStr(rsxls("�γ̱��"))+"','"+CStr(rsxls("�γ�����"))+"',"+CStr(rsxls("ѧʱ"))+","+CStr(rsxls("ѧ��"))++",'"+CStr(rsxls("��ʦ1"))+"','"+CStr(teacher2)+"','"+CStr(rsxls("�γ�����"))+"','"+savedfile+"','"+CStr(rsxls("����ʱ��"))+"', "+cstr(rsxls("����"))+",0)"
					'strsql3 = "insert [coursemember] (stuid,stuname,stuobject,stucollege,courseid,coursename,isevaluate) values ('"+CStr(rsxls("ѧ��"))+"','"+CStr(rsxls("����"))+"','"+CStr(rsxls("רҵ"))+"','"+CStr(rsxls("�Ͽ�ϵ"))+"','"+courseid+"','"+coursename+"',0)"
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
				'response.write "<script>alert('���棡');history.back();</script>"' û��ִ��
			End If 
			
			'ѡ������
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
				response.write "<script>alert('�ϴ�ѡ�������ļ�û����Ч���ݣ�');window.location.href='adminupload.asp';</script>"
				responser.end
				Else 
			
			'===============
			'���ݿ���ʱ����ɾ��[coursemember]�����еĿγ̣��Ӷ���֤Ψһ��
			strsql = "delete from [coursemember] where coursedate ='"+cstr(rsxls("����ʱ��"))+"'"
			conn.execute strsql
			
			'===============								
				Do While Not rsxls.eof
					strsql = "insert into [coursemember] (stuid,stuname,stuobject,stucollege,courseid,coursedate,courseorder,isevaluate) values ('"+CStr(rsxls("ѧ��"))+"','"+CStr(rsxls("����"))+"','"+CStr(rsxls("רҵ"))+"','"+CStr(rsxls("�Ͽ�ϵ"))+"','"+CStr(rsxls("�γ̺�"))+"','"+CStr(rsxls("����ʱ��"))+"',"+cstr(rsxls("����"))+",0)"
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
					response.write "<script>alert('upload member success��');</script>"
				Else 
					conn.RollbackTrans
					response.write "<script>alert('upload member failed��');</script>"
				End If 
				'conn.close()
				'Set conn = Nothing 
				
			End If 	
			response.write "<script>alert('����ȡ�� "+cstr(totalcourse)+" �ſ�, "+cstr(totalmember)+" ���Ͽμ�¼, ��ȷ�������Ƿ���Ͻ��');</script>"
			'�������ݱ�
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
