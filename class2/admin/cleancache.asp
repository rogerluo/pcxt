<%
'-------------------------------------------
' Dvbbs System Update Software Tools
' ClearCache File 
' �����ȷ� [AspSky Software, Inc.]
' ScriptEditor Fssunwin
' 2005-03-25
'-------------------------------------------
Call RemoveAllCache()

Sub RemoveAllCache()
Dim cachelist,i
Call InnerHtml("UpdateInfo","<b>��ʼִ��������ǰվ�㻺��</b>��")
Cachelist=split(GetallCache(),",")
If UBound(cachelist)>1 Then
For i=0 to UBound(cachelist)-1
DelCahe Cachelist(i)
Call InnerHtml("UpdateInfo","���� <b>"&cachelist(i)&"</b> ���")
Next
Call InnerHtml("UpdateInfo","������"& UBound(cachelist)-1 &"���������<br>")
Else
Call InnerHtml("UpdateInfo","<b>��ǰվ��ȫ������������ɡ�</b>��")
End If
End Sub

Function GetallCache()
Dim Cacheobj
For Each Cacheobj in Application.Contents
GetallCache = GetallCache & Cacheobj & ","
Next
End Function

Sub DelCahe(MyCaheName)
Application.Lock
Application.Contents.Remove(MyCaheName)
Application.unLock
End Sub

Sub InnerHtml(obj,msg)
Response.Write "<li>"&msg&"</li>"
Response.Flush
End Sub
%>