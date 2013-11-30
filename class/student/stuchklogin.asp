<!--#include file="../admin/conn.asp"-->
<%
	stuname=replace(trim(request.form("stuname")),"'","")
	stuid=(trim(request.form("stuid")))

	'response.write stuid+"<p>"
	'response.write adminname

	'if stuname = "" Or stuid="" Or (Not IsNumeric(stuid)) then
	if stuname = "" Or stuid="" then
		conn.close
		set conn = nothing
		response.Redirect "stulogin.asp"
		response.end
	end if

	strsql = "select stuid from [coursemember]  where stuname='"&stuname&"' and stuid='"&stuid&"'"
	set rs = conn.execute(strsql)

	if Not (rs.eof And rs.bof) then
		session("stuname") = stuname
		session("stuid") = stuid
		session.Timeout=30
		set rs=nothing
		conn.close
		set conn = nothing
		response.Redirect "stumain.asp"
	else
		response.write "<script language='javascript'>alert('ĞÕÃû»òÑ§ºÅ´íÎó£¡');history.go(-1);</script>"
		set rs=nothing
		conn.close
		set conn = nothing
	End If
%>