<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	target = Trim(request.querystring("target"))
	 
	strsql = "select [evaluate2_gra].stuid,[evaluate2_gra].stuname,[evaluate2_gra].coursestudy, [evaluate2_gra].coursejob from [coursemember_gra] join [evaluate2_gra] on ([coursemember_gra].stuid = [evaluate2_gra].stuid and [coursemember_gra].courseid = '"+courseid+"' and [coursemember_gra].coursedate = '"+coursedate+"')"	
	Set rs = conn.execute(strsql)

%>
<table width="350" border="1" align="center" cellpadding="5" cellspacing="0" class="tbbg">
	<tr bgcolor=#cccccc>
<%
	If target = 0 Then 
%>
		<th colspan=2>
			支持课程对于自身深造有帮助的学生
		</th>
<%
	Else
%>
		<th colspan=2>
			支持课程对于自身找工作有帮助的学生
		</th>	
<%
	End If 
%>
	</tr>
<%
	If rs.eof Or rs.bof Then 
	Else
		i = 0
		Do While Not rs.eof
			If i Mod 2 = 0 Then
				response.write "<tr>"
			Else 
				response.write "<tr bgcolor=#cccccc>"
			End If
			If target = 0 Then 
				str = rs("coursestudy")
			Else
				str = rs("coursejob")
			End If

			strarray = Split(str,",")

			For j = 0 To UBound(strarray)
				If Trim(strarray(j)) = (courseid) Then
					response.write "<td width='50%'>"
					response.write rs("stuid")
					response.write "</td><td>"
					response.write rs("stuname")
					response.write "</td></tr>"
					i = i + 1
					Exit For
				End If 
			Next
			rs.movenext
		Loop 
	End If 
%>
</table>