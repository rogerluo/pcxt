<%
'-------------------------------------------
' Dvbbs System Update Software Tools
' ClearCache File 
' 动网先锋 [AspSky Software, Inc.]
' ScriptEditor Fssunwin
' 2005-03-25
'-------------------------------------------
Call RemoveAllCache()

Sub RemoveAllCache()
Dim cachelist,i
Call InnerHtml("UpdateInfo","<b>开始执行清理当前站点缓存</b>：")
Cachelist=split(GetallCache(),",")
If UBound(cachelist)>1 Then
For i=0 to UBound(cachelist)-1
DelCahe Cachelist(i)
Call InnerHtml("UpdateInfo","更新 <b>"&cachelist(i)&"</b> 完成")
Next
Call InnerHtml("UpdateInfo","更新了"& UBound(cachelist)-1 &"个缓存对象<br>")
Else
Call InnerHtml("UpdateInfo","<b>当前站点全部缓存清理完成。</b>。")
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
