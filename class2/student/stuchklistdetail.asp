<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');history.back();</script>"
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
		response.write "<script>alert('传输信息有误！');window.location.href='stuevaluatedetail?courseid="+courseid+".asp';</script>"
		response.end
	End If 

	strsql = "update [evaluate_gra] set coursecontent='"+coursecontent+"',coursedifficult='"+coursedifficult+"',coursepoint='"+coursepoint+"',courselang='"+courselang+"',courseknowhelp='"+courseknowhelp+"',coursejobhelp='"+coursejobhelp+"',coursecontinue='"+coursecontinue+"'"
	
	conn.BeginTrans
	conn.execute(strsql)
	If conn.Errors.Count = 0 Then
		conn.CommitTrans
		response.write "<script>alert('更新成功！');window.location.href='stulist.asp';</script>"
	Else 
		conn.RollbackTrans
		response.write "<script>alert('更新失败！');window.location.href='stulist.asp';</script>"
	End If
	
	conn.close()
	Set conn = Nothing
%>