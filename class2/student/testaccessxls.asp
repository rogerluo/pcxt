<!--#include file="../admin/conn.asp"-->
<%
dim connxls,connxlsstr
savedfile = "E:\pcxt\class\admin\UploadFiles\myplan.xls"
connxlsstr ="Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;HDR=YES;IMEX=1; Data Source="+savedfile+" 
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
			strsqlxls = "select * from [Sheet2$]"
			Set rsxls = connxls.execute(strsqlxls)
			'��[coursemember]����'
			'Set conn3 = server.CreateObject("ADODB.Connection")
			'conn3.open connstr
			'conn3.BeginTrans
			conn.BeginTrans
			If rsxls.eof And rsxls.bof Then
				connxls.close()
				rsxls.close()
				'conn3.close()
				'conn.close()
				'Set conn = Nothing 
				response.write "<script>alert('�ϴ��ļ�û����Ч���ݣ�');window.location.href='adminupload.asp';</script>"
				responser.end
			Else 
				Do While Not rsxls.eof
					response.write CSTr(rsxls(0)) +"<p>"
					'strsql = "insert into [courseinfo] (courseid) values('"+CStr(rsxls(0))+"')"
					teacher2 = rsxls("��ʦ2")
					If IsNull(teacher2) Then 
						teacher2 = ""
					End If 
					
					strsql = "insert into [courseinfo] (courseid,coursename,coursetime,coursecredit,courseteacher,courseteacher2,coursetype,coursefile,courseheadcount) values ('"+CStr(rsxls("�γ̱��"))+"','"+CStr(rsxls("�γ�����"))+"',"+CStr(rsxls("ѧʱ"))+","+CStr(rsxls("ѧ��"))++",'"+CStr(rsxls("��ʦ1"))+"','"+teacher2+"','"+CStr(rsxls("�γ�����"))+"','"+savedfile+"',0)"
					'strsql3 = "insert [coursemember] (stuid,stuname,stuobject,stucollege,courseid,coursename,isevaluate) values ('"+CStr(rsxls("ѧ��"))+"','"+CStr(rsxls("����"))+"','"+CStr(rsxls("רҵ"))+"','"+CStr(rsxls("�Ͽ�ϵ"))+"','"+courseid+"','"+coursename+"',0)"
					'Set rs3 = conn3.execute(strsql3)
					conn.execute strsql
					rsxls.movenext
				Loop
				rsxls.close()
				Set rsxls = Nothing 
				connxls.close()
				Set connxls = Nothing 
				If conn.Errors.Count = 0 Then
					conn.CommitTrans
					response.write("Update [courseinfo] success<P>")
				Else 
					conn.RollbackTrans
					response.write("Update [courseinfo] failed<p>")
				End If 
				conn.close()
				Set conn = Nothing 

			End If 
%>