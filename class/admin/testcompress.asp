<!--#include file="compress.asp"-->
<%
Sub Debug(vData)
	response.write vData+"<br>"
End Sub
	Set ObjCompress = new COMPRESS_DECOMPRESS_FILES
	ObjCompress.compressPath="学期课程评价"
	ObjCompress.decompresspath="ExcelFile"
	ObjCompress.compress
	mFileName = Server.Mappath("学期课程评价.rar")
	'判断文件是否存在
	Set ObjFso=Server.CreateObject("Scripting.FileSystemObject")
	If ObjFso.FileExists(mFileName) then
		Debug "Exists"
	Else 
		Debug "Not Exists"
	End If
	Set ObjFso = Nothing
	Set ObjCompress = Nothing 
%>