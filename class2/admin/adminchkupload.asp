<!-- #include file="Upload.asp" -->
<!--#include file="conn.asp"-->
<!--#include file="chkstate.asp"-->
<%
'============================================
	If session("sysstate") = 1 Then 
		conn.close()
		set conn = Nothing
		response.write "<script>alert('ϵͳ��û�رգ�');window.location.href='adminmain.asp';</script>"
		response.end
	End If 
'============================================
		'��ȡ�ļ���������������׺��
		Function GetFileName(ByVal strFile)
			If strFile <> "" Then
				startnum = InstrRev(strFile,"\")+1
				endnum = InstrRev(strFile,".")
				GetFileName = mid(strFile,startnum,endnum-startnum)
			Else
				GetFileName = ""
			End If
		End  Function
		'��ȡ�ϴ��ļ��ļ����ͣ����ļ���׺��
		Function GetFileType(ByVal strFile)
			If strFile <> "" Then 
				GetFileType = Mid(strFile,InstrRev(strFile,".")+1)
			Else 
				GetFileType = ""
			End If 
		End Function
		'��ȡ�ϴ��ļ��ļ�ȫ���������ļ���׺��
		Function GetFileFullName(ByVal strFile)
			If strFile <> "" Then
				GetFileFullName = mid(strFile,InStrRev(strFile, "\")+1)
			Else
				GetFileFullName = ""
			End If
		End  Function
'============================================
'����form���漰��excel�ļ�
		formPath = "UploadFile/"
		Set upload= New DoteyUpload
		Upload.SaveTo(formPath) '���ļ��������ļ���ͳһ������ĳ·����
		If upload.ErrMsg = "" then
			response.write "<script>alert('upload sucess!');</script>"
		Else 
			response.write "<script>alert('upload failed!');</script>"	
		End If 
		Set upload=Nothing
'===========================================
		'�γ̼ƻ�+�γ�����
		courseplan = Server.MapPath("./UploadFile") + "\courseplan.xls"
		coursemember = Server.MapPath("./UploadFile") + "\coursemember.xls"

		dim connxls,connxlsstr
	
		connxlsstr ="Provider='Microsoft.Jet.OLEDB.4.0';Extended Properties='Excel 8.0;HDR=YES;IMEX=1';Data Source="+courseplan+";"
	
		Set connxls=Server.CreateObject("ADODB.Connection") 
		Set rsxls = server.CreateObject("ADODB.Recordset")

		connxls.open connxlsstr
		strsqlxls = "select * from [Sheet1$]"
		Set rsxls = connxls.execute(strsqlxls)
		'������	
		conn.BeginTrans
		
		If rsxls.eof And rsxls.bof Then
			connxls.close()
			rsxls.close()
			response.write "<script>alert('�ϴ��γ̼ƻ��ļ�û����Ч���ݣ�');window.location.href='adminupload.asp';</script>"
			responser.end
		Else 	
			'��ɾ�����ݿ���ӵ����ͬ����ʱ��Ŀγ�
			strsql = "delete from [courseinfo_gra] where coursedate ='"+cstr(rsxls("����ʱ��"))+"'"
			conn.execute strsql
								
			Do While Not rsxls.eof
			
				strsql = "insert into [courseinfo_gra] (courseid,coursename,coursetime,coursecredit,courseteacher,coursetype,coursetype2,coursedate) values ('"+CStr(rsxls("�γ̱��"))+"','"+CStr(rsxls("�γ�����"))+"',"+(CStr(rsxls("ѧʱ")))+","+(CStr(rsxls("ѧ��")))+",'"+CStr(rsxls("�ον�ʦ"))+"','"+CStr(rsxls("�γ����"))+"','"+CStr(rsxls("���"))+"','"+CStr(rsxls("����ʱ��"))+"')"
					
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
			End If'�ж������Ƿ�ִ�гɹ�
		End If'�ж��ļ��Ƿ�Ϊ��
			
		'ѡ������
		connxlsstr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+coursemember+";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';" 
		Set connxls=Server.CreateObject("ADODB.Connection") 
		Set rsxls = server.CreateObject("ADODB.Recordset")

		connxls.open connxlsstr
		strsqlxls = "select * from [Sheet1$]"
		Set rsxls = connxls.execute(strsqlxls)

		'������	
		conn.BeginTrans
		If rsxls.eof And rsxls.bof Then
			connxls.close()
			rsxls.close()
			response.write "<script>alert('�ϴ�ѡ�������ļ�û����Ч���ݣ�');window.location.href='adminupload.asp';</script>"
			responser.end
		Else 
			'��ɾ�����ݿ���ӵ����ͬ����ʱ���ѡ������
			strsql = "delete from [coursemember_gra] where coursedate ='"+cstr(rsxls("����ʱ��"))+"'"
			conn.execute strsql

			Do While Not rsxls.eof
				strsql = "insert into [coursemember_gra] (stuid,stuname,stuobject,stucollege,courseid,coursedate,isevaluate) values ('"+CStr(rsxls("ѧ��"))+"','"+CStr(rsxls("����"))+"','"+CStr(rsxls("רҵ"))+"','"+CStr(rsxls("�Ͽ�ϵ"))+"','"+CStr(rsxls("�γ̺�"))+"','"+CStr(rsxls("����ʱ��"))+"',0)"

				conn.execute strsql
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
			End If'�ж������Ƿ�ִ�гɹ� 
		End If'�ж��ļ��Ƿ�Ϊ��	
		
		'�������ݱ�
		conn.begintrans
		'ͬ��ѡ�������пγ�����	
		strsql = "update [coursemember_gra] set [coursemember_gra].coursename = [courseinfo_gra].coursename from [coursemember_gra] left join [courseinfo_gra] on ([coursemember_gra].courseid = [courseinfo_gra].courseid)"
			
		conn.execute strsql
		
		'ͬ��ѡ�������пγ�����
		strsql = "update [courseinfo_gra] set [courseinfo_gra].courseheadcount = (select count([coursemember_gra].stuid) from [coursemember_gra] where [coursemember_gra].courseid = [courseinfo_gra].courseid) from [courseinfo_gra]"		
			
		conn.execute strsql

		'ͬ��ѡ�������пγ̵����
		strsql = "update [coursemember_gra] set [coursemember_gra].isevaluate = 1 from [coursemember_gra] join [evaluate_gra] on ([coursemember_gra].stuid = [evaluate_gra].stuid and [coursemember_gra].courseid = [evaluate_gra].courseid)"

		conn.execute strsql
		
		if conn.Errors.Count = 0 then 
			conn.committrans
			conn.close()
			set conn = nothing
			response.write "<script>history.back();</script>"
		End if	
%>