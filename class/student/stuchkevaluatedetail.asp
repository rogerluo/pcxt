<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	courseorder = Trim(request.querystring("courseorder"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" or courseorder = "" or Not IsNumeric(courseorder) Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');history.back();</script>"
		response.end
	End If 	
	
	stuid = session("stuid")
	stuname = session("stuname")


	book = Trim(request.Form("book"))
	courseraw = Trim(request.Form("courseraw"))
	rawlang = Trim(request.Form("rawlang"))
	lang = Trim(request.Form("lang"))
	scale = Trim(request.Form("scale"))
	spirit = Trim(request.Form("spirit"))
	info = Trim(request.Form("info"))
	teachstatus = Trim(request.Form("teachstatus"))
	teachtime = Trim(request.Form("teachtime"))
	teachtime2 = Trim(request.Form("teachtime2"))
	teachername2 = Replace(Trim(request.Form("teachername2")),"'","")
	stoptimes = Trim(request.Form("stoptimes"))
	stopcourse = Replace(Trim(request.Form("stopcourse")),"'","")
	retimes = Trim(request.Form("retimes"))
	total = Trim(request.Form("total"))
	advice = Replace(Trim(request.Form("advice")),"'","")
	
	If book = "" Or courseraw = "" Or rawlang = "" Or lang = "" Or scale = "" Or spirit = "" Or info = "" Or teachstatus = "" Or teachtime = "" Or total = "" Then
		conn.close()
		Set conn = Nothing 
		response.write "<script>alert('传输信息有误！');history.back();</script>"
		response.end
	End If 

	strsql = "delete from [evaluate] where courseid = '"+courseid+"' and coursedate = '"+coursedate+"' and stuid = '"+stuid+"' and courseorder ="+courseorder
	conn.execute(strsql)

	strsql = "insert  [evaluate] (courseid,coursedate, courseorder,stuid,stuname,book,courseraw,rawlang,lang,scale,spirit,info,teachstatus,teachtime,teachtime2,teachername2,stoptimes,stopcourse,retimes,total,advice) values('"+courseid+"','"+coursedate+"',"+courseorder+",'"+stuid+"','"+stuname+"','"+book+"','"+courseraw+"','"+rawlang+"','"+lang+"','"+scale+"','"+spirit+"','"+info+"','"+teachstatus+"','"+teachtime+"','"+teachtime2+"','"+teachername2+"','"+stoptimes+"','"+stopcourse+"','"+retimes+"','"+total+"','"+advice+"')"
	
	conn.BeginTrans
	conn.execute(strsql)
	If conn.Errors.Count = 0 Then
		conn.CommitTrans
		'response.write("Insert [evaluate] success<P>")
	Else 
		conn.RollbackTrans
		'response.write("Insert [evaluate] failed<p>")
	End If
	
	strsql = "update [coursemember] set isevaluate=1 where stuid='"+stuid+"' and courseid ='"+courseid+"' and courseorder="+courseorder
	conn.execute(strsql)

	conn.close()
	Set conn = Nothing

	response.write "<script>alert('更新数据库成功！');window.location.href='stuevaluate.asp';</script>"
%>