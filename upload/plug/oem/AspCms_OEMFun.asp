<!--#include file="../../inc/AspCms_MainClass.asp" -->
<%
CheckAdmin("AspCms_OEM.asp")
dim adminpath
adminpath="admin"


'On Error Resume Next
dim action : action=getform("action","get")
Dim objUpdate,updatelist,newver,info,outmes,vnum,vstr,fs,username,password
if action="down" then

	
	if GetBody(getForm("username","post"),getForm("password","post"))="-1" then
		alertMsgAndGo "用户名或密码错误!","-1"
	elseif GetBody(getForm("username","post"),getForm("password","post"))="-2" then
		alertMsgAndGo "对不起，您的账号已被禁用!","-1"
	else
		Set objUpdate = New Cls_oUpdate
			With objUpdate
			  .action=action
			  .UrlVersion = "http://oem.aspcms.com/"
			  .UrlUpdate = "http://oem.aspcms.com/"
			  .UpdateLocalPath = "/"
			  .LocalVersion =vnum
			  .FileType      = ".asp"
			  .doUpdateVersion(GetBody(getForm("username","post"),getForm("password","post")))
			  outmes=.outmes
			  info= .Info
			  updatelist = .updatelist
			  newver=.newver
			End With
		Set objUpdate = Nothing
	end if
	
end if


Function GetBody(username,password)
dim https
Set https = Server.CreateObject("MSXML2.XMLHTTP") 
With https 
.Open "Post", "http://oem.aspcms.com/app.asp", False
.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
.Send "LoginName="& username &"&userPass="&password
GetBody = .ResponseBody
End With 
GetBody = BytesToBstr(GetBody,"GB2312")
Set https = Nothing 
End Function

Function BytesToBstr(body,Cset) 
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText 
objstream.Close
set objstream = nothing
End Function

Class Cls_oUpdate
  Public LocalVersion, LastVersion, FileType,action,newver,updatelist,updatever
  Public UrlVersion, UrlUpdate, UpdateLocalPath, Info,outmes,yesup,newsitepath
  Public UrlHistory
  Private sstrVersionList, sarrVersionList, sintLocalVersion, sstrLocalVersion
  Private sstrLogContent, sstrHistoryContent, sstrUrlUpdate, sstrUrlLocal



   Public function doUpdateVersion(oemkey)
   		Call doUpdateFile(UrlUpdate&"oemuser/" & oemkey&"/login.txt", "../../"&adminpath&"/login.asp")
		Call doUpdateFile(UrlUpdate&"oemuser/" & oemkey&"/menu.txt", "../../"&adminpath&"/menu.asp")
		Call doUpdateFile(UrlUpdate&"oemuser/" & oemkey&"/menu_user.txt", "../../"&adminpath&"/menu_user.asp")
		call saveimage(UrlUpdate&"oemuser/" & oemkey&"/logo.jpg", "../../"&adminpath&"/images/admin_tlogo.jpg")
		call saveimage(UrlUpdate&"oemuser/" & oemkey&"/logo.jpg", "../../"&adminpath&"/images/logo_user.jpg")
		call saveimage(UrlUpdate&"oemuser/" & oemkey&"/loginlogo.jpg", "../../"&adminpath&"/images/logo_01.gif")
		Call doUpdateFile(UrlUpdate&"oemuser/" & oemkey&"/AspCms_SettingClass.txt", "../../"&adminpath&"/inc/AspCms_SettingClass.asp")
		
		
		alertMsgAndGo "信息更新成功","Aspcms_OEM.asp"
   End function
'******************************************************************
   Private function doUpdateFile(strSourceFile, strTargetFile)
    Dim strContent
    strContent = GetContent(sstrUrlUpdate & strSourceFile)

    If sDoCreateFile(Server.MapPath(sstrUrlLocal &strTargetFile), strContent) Then     
     	sstrLogContent = sstrLogContent & "  更新成功" & vbCrLf
    Else
     	sstrLogContent = sstrLogContent & "  更新失败" & vbCrLf
    End If
   End function
'************************************************************

	function saveimage(from,tofile) 
		dim url,savepath,filename,xmlhttp,img,text,objAdostream
		url=from 
		savepath=tofile 
		set xmlhttp=server.createobject("Microsoft.XMLHTTP") 
		xmlhttp.open "get",url,false 
		xmlhttp.send 
		img = xmlhttp.ResponseBody
		'text=xmlhttp.ResponseText
		set xmlhttp=nothing 
		set objAdostream=server.createobject("ADODB.Stream") 
		objAdostream.Open() 
		objAdostream.type=1 
		objAdostream.Write(img) 
		objAdostream.SaveToFile server.MapPath(tofile),2  
		objAdostream.SetEOS 
		set objAdostream=nothing 
	end function 


   Private function GetContent(strUrl)
    GetContent = ""
    
    Dim oXhttp, strContent
    Set oXhttp = Server.CreateObject("Microsoft.XMLHTTP")
    On Error Resume Next 
    With oXhttp
     .Open "GET", strUrl,False, "", ""
     .Send
     If .readystate <> 4 Then Exit function
     strContent = .Responsebody
     strContent = sBytesToBstr(strContent)
    End With
    
    Set oXhttp = Nothing
    If Err.Number <> 0 Then
     response.Write(Err.Description)
     Err.Clear
     Exit function
    End If
    
    GetContent = strContent
   End function

   Private function sBytesToBstr(vIn)
    dim objStream
    set objStream = Server.CreateObject("adodb.stream")
    objStream.Type    = 1
    objStream.Mode    = 3
    objStream.Open
    objStream.Write vIn
    
    objStream.Position  = 0
    objStream.Type    = 2
    objStream.Charset  = "GB2312"
    sBytesToBstr     = objStream.ReadText 
    objStream.Close
    set objStream    = nothing
   End function

   Private function sDoCreateFile(strFileName, ByRef strContent)
    sDoCreateFile = False
    Dim strPath
    strPath = Left(strFileName, InstrRev(strFileName, "\", -1, 1))

   If Not(CreateDir(strPath)) Then Exit function
    'If Not(CheckFileName(strFileName)) Then Exit function
    
    'response.Write(strFileName)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(strFileName, ForWriting, True)
    f.Write strContent
    f.Close
    Set fso = nothing
    Set f = nothing
    sDoCreateFile = True
   End function

   Private function sDoAppendFile(strFileName, ByRef strContent)
    sDoAppendFile = False
    Dim strPath
    strPath = Left(strFileName, InstrRev(strFileName, "\", -1, 1))

    If Not(CreateDir(strPath)) Then Exit function

    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
	
	
    Set f = fso.OpenTextFile(strFileName, ForAppending, True)
    f.Write strContent
    f.Close
    Set fso = nothing
    Set f = nothing
    sDoAppendFile = True
   End function

   Private function CreateDir(ByVal strLocalPath)
    Dim i, strPath, objFolder, tmpPath, tmptPath
    Dim arrPathList, intLevel
    
    On Error Resume Next
    strPath     = Replace(strLocalPath, "\", "/")
    Set objFolder  = server.CreateObject("Scripting.FileSystemObject")
    arrPathList   = Split(strPath, "/")
    intLevel     = UBound(arrPathList)
    
    For I = 0 To intLevel
     If I = 0 Then
      tmptPath = arrPathList(0) & "/"
     Else
      tmptPath = tmptPath & arrPathList(I) & "/"
     End If
     tmpPath = Left(tmptPath, Len(tmptPath) - 1)
     If Not objFolder.FolderExists(tmpPath) Then objFolder.CreateFolder tmpPath
    Next
    
    Set objFolder = Nothing
    If Err.Number <> 0 Then
     CreateDir = False
     Err.Clear
    Else
     CreateDir = True
    End If
   End function

   Private function toNum(s, default)
    If IsNumeric(s) and s <> "" then
     toNum = CLng(s)
    Else
     toNum = default
    End If
   End function
 End Class
  If Err.Number <> 0 Then
     echo err.description
  	Err.Clear
  end if
%>
