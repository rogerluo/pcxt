<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('����ҳ�����');history.back();</script>"
		response.end
	End If 	
	
	stuid = session("stuid_gra")
	stuname = session("stuname_gra")


	coursecontent = Trim(request.Form("coursecontent"))
	coursedifficult = Trim(request.Form("coursedifficult"))
	coursepoint = Trim(request.Form("coursepoint"))
	courselang = Trim(request.Form("courselang"))
	courseknowhelp = Trim(request.Form("courseknowhelp"))
	coursejobhelp = Trim(request.Form("coursejobhelp"))
	coursecontinue = Trim(request.Form("coursecontinue"))
	
	If coursecontent = "" Or coursedifficult = "" Or coursepoint = "" Or courselang = "" Or courseknowhelp = "" Or coursejobhelp = "" Or coursecontinue = "" Then
		conn.close()
		Set conn = Nothing 
		response.write "<script>alert('������Ϣ����');history.back();</script>"
		response.end
	End If 


	strsql = "insert  [evaluate_gra] (courseid,coursedate,stuid,stuname,coursecontent,coursedifficult,coursepoint,courselang,courseknowhelp,coursejobhelp,coursecontinue) values('"+courseid+"','"+coursedate+"','"+stuid+"','"+stuname+"','"+coursecontent+"','"+coursedifficult+"','"+coursepoint+"','"+courselang+"','"+courseknowhelp+"','"+coursejobhelp+"','"+coursecontinue+"')"
	
	conn.BeginTrans
	conn.execute(strsql)
	If conn.Errors.Count = 0 Then
		conn.CommitTrans
	Else 
		conn.RollbackTrans
		response.write "<script>alert('�������ݿ�ʧ�ܣ�');window.location.href='stuevaluate.asp';</script>"
	End If
	
	strsql = "update [coursemember_gra] set isevaluate=1 where stuid='"+stuid+"' and courseid ='"+courseid+"'"
	conn.execute(strsql)

	conn.close()
	Set conn = Nothing

	response.write "<script>alert('�������ݿ�ɹ���');window.location.href='stuevaluate.asp';</script>"
%>