<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<%
	strsql = "delete from evaluate2_gra where stuid = '"+session("stuid_gra")+"'"
	conn.execute(strsql)

	strsql = "insert into evaluate2_gra (stuid,stuname,coursestudy,coursejob) values ('"+session("stuid_gra")+"','"+session("stuname_gra")+"','"+request.Form("helpstudy")+"','"+request.Form("helpjob")+"')"

	conn.execute(strsql)

	response.write "<script>alert('更新成功!');history.back();</script>"
%>