<!--#include file="conn.asp"-->
<!--#include file="Cls\ExcelControl.asp"-->
<!--#include file="compress.asp"-->
<%
	'删除ExcelFile文件夹（删除里面内容）
	strFolder = "ExcelFile"
	Call DeleteFolder(strFolder)
	Call CreateFolder(strFolder)

	coursedate = Trim(request.querystring("coursedate"))
	strsql = "select * from courseinfo where coursedate = '"+coursedate+"'"
	Set rs = conn.execute(strsql)
	'判断记录集是否为空
	If rs.eof Or rs.bof Then 
		rs.close
		Set rs = Nothing 
		response.End 
	End If
	'给各个课程创建对应评价文件
	'Set ObjExl=New ExcelMarker
	Do While Not rs.eof
		Set ObjExl=New ExcelMarker
		ObjExl.SetCourseid = CStr(rs("courseid"))
		ObjExl.SetCourseDate = CStr(rs("coursedate"))
		ObjExl.SetCourseOrder = CStr(rs("courseorder"))
		ObjExl.SetItemArray
		
		Set AdoRs=Server.CreateObject("ADODB.RecordSet")
		AdoRs.Open"select * From totalevaluate",conn,1,3

		ObjExl.ADOStreamExcel AdoRs

		rs.movenext
	Loop
	rs.close
	Set rs=Nothing

	'提前删除压缩文件
	Set ObjFso=Server.CreateObject("Scripting.FileSystemObject")
	mFileName = Server.Mappath(coursedate+"学期课程评价.rar")
	'判断文件是否存在
	If ObjFso.FileExists(mFileName) Then
		Debug "发现已有文件并删除"
		Debug mFileName
		ObjFso.DeleteFile(mFileName)  
	End If
	Set ObjFso = Nothing

	'压缩ExcelFile文件夹里文件
	Set ObjCompress = new COMPRESS_DECOMPRESS_FILES
	ObjCompress.compressPath=coursedate+"学期课程评价"
	ObjCompress.decompresspath="ExcelFile"
	ObjCompress.compress

	Function DeleteFolder(strFolder) 
	   '首选判断要建立的文件夹是否已经存在 
		Dim strTestFolder, objFSO 
		strTestFolder = Server.Mappath(strFolder) 
		Set objFSO = CreateObject("Scripting.FileSystemObject") 
		' 检查文件夹是否存在 
		If objFSO.FolderExists(strTestFolder) Then 
			objFSO.DeleteFolder(strTestFolder) 
			Debug "删除成功！"
		Else 
			Debug "该文件夹不存在！"
		End If 
		Set objFSO = Nothing 
	End Function 
	Function CreateFolder(strFolder)
		set objFSO = server.CreateObject("scripting.filesystemobject")
		objFSO.createfolder(server.MapPath(strFolder))
		Debug "创建成功！"
	End Function 

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>电子信息工程学院研究生教学状况调查系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
	response.write "下载连接"
	response.write "<a href='"+coursedate+"学期课程评价.rar'>Download</a>"
%>
</body>
</html>