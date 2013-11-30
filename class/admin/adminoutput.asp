<!--#include file="conn.asp"-->
<!--#include file="Cls\ExcelControl.asp"-->
<%
Function download(FilePath,FileName) 
	'FilePath文件全路径，FileName下载文件的文件名 
    On Error Resume Next
	Set FSObj = server.CreateObject("scripting.filesystemobject")
	Set fn=FSObj.GetFile(FilePath)
	FileSize=fn.size
	Set fn = Nothing

    Set AdoObj=CreateObject("Adodb.Stream") 
    AdoObj.Open 
    AdoObj.Type=1 
    AdoObj.LoadFromFile(FilePath) 
   
	Response.ContentType="application/vnd.ms-excel" 
	'Response.ContentType="application/x-zip-compressed"
	Response.AddHeader "Content-Disposition","Attachment;filename="&FileName
	Response.AddHeader "Content-Length",FileSize
	Response.CharSet="UTF-8"
	Response.BinaryWrite AdoObj.Read  
	Response.Flush   
	Response.Clear() 
	AdoObj.Close
	Set AdoObj=Nothing 
End Function
%>
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	courseorder = Trim(request.querystring("courseorder"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" or courseorder = "" or Not IsNumeric(courseorder) Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');window.location.href='adminlist.asp';</script>"
		response.end
	End If
	
	'导出excel文件
	Set ObjExl=New ExcelMarker
	Set AdoRs=Server.CreateObject("ADODB.RecordSet")
	AdoRs.Open"select * From totalevaluate",conn,1,3

	ObjExl.SetCourseid = courseid
	ObjExl.SetCourseDate = coursedate
	ObjExl.SetCourseOrder = courseorder

	ObjExl.ADOStreamExcel AdoRs'ADODB.RecordSet对象
	
	mFilePath=ObjExl.FileName
	savefilename = ObjExl.GetCourseName+".xls"

	AdoRs.Close
	conn.Close
	Set AdoRs=Nothing
	Set conn=Nothing



	'下载excel文件
	Call download(mFilePath,savefilename)
%>