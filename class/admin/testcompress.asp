<!--#include file="compress.asp"-->
<%
Sub Debug(vData)
	response.write vData+"<br>"
End Sub
	Set ObjCompress = new COMPRESS_DECOMPRESS_FILES
	ObjCompress.compressPath="ѧ�ڿγ�����"
	ObjCompress.decompresspath="ExcelFile"
	ObjCompress.compress
	mFileName = Server.Mappath("ѧ�ڿγ�����.rar")
	'�ж��ļ��Ƿ����
	Set ObjFso=Server.CreateObject("Scripting.FileSystemObject")
	If ObjFso.FileExists(mFileName) then
		Debug "Exists"
	Else 
		Debug "Not Exists"
	End If
	Set ObjFso = Nothing
	Set ObjCompress = Nothing 
%>