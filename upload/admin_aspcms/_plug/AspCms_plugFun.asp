<!--#include file="../inc/AspCms_SettingClass.asp" -->

<%CheckAdmin("AspCms_plug.asp")%>
<%
dim action : action=getForm("action","get")
dim acttype : acttype=getform("acttype", "get")
dim filename,path,filetext,nametype

path="../plug"

select case action
	case "add" : plugadd
	case "del" : plugDel
end select

Sub getFile()
	filename=getForm("filename","get")	
	filetext=loadFile(path&filename)
End Sub

Sub MsgBoxNGo(msg,url)
dim script
'script = "<" & "script>"
script = script & msg
script = script & "<a href='"&url&"' onclick='window.parent.document.frames[0].location.reload();'>点击返回</a>"
'script = script & "alert(window.parent.document.frames[0].document.location);"
'script = script & ".location='"&url&"';"
'script = script & "</" & "script>"
response.write script
response.end
End Sub

Sub plugadd
	dim plugkey,Menuname,MenuLink,plugArray,sql,Menunames,MenuLinks,i,plugpath
	plugkey=getForm("plugkey","get")
	Menuname=getForm("Menuname","get")
	MenuLink=getForm("MenuLink","get")
	plugpath=getForm("plugpath","get")
	if isnul(Menuname) then alertMsgAndGo"安装出错","-1"
	if isnul(plugkey) then alertMsgAndGo"安装出错","-1"
	if isnul(MenuLink) then alertMsgAndGo"安装出错","-1"
	plugArray=conn.Exec("select * from {prefix}Menu where MenuKey='"&plugkey&"'","arr")
	if isarray(plugArray) then
		if isnul(plugname) then alertMsgAndGo"已安装的模块或重复的插件标识","-1"
	else
	Menunames=split(Menuname,"|")
	MenuLinks=split(MenuLink,"|")
	for i=0 to ubound(Menunames)
	sql="insert into {prefix}Menu(ParentID, TopID, MenuName, MenuLink, MenuOrder, MenuLevel, MenuStatus,Menukey) values(67, 67, '"& Menunames(i)&"', '"&path & "/"& plugpath&"/"&MenuLinks(i)&"', 1, 2, 1,'"&plugkey&"')"
	conn.exec sql,"exe"
	next
	if session("groupMenu")<>"all" then
	Dim rs :set rs =Conn.Exec ("select Menuid from {prefix}Menu where menukey='"& plugkey &"'","r1")
	IF rs.eof or rs.bof Then
		alertMsgAndGo"安装出错","-1"
	else
		Do While not rs.eof 
			sql="update {prefix}UserGroup set GroupMenu=GroupMenu & ', "& rs("Menuid") &"' where Groupid in (select Groupid from {prefix}user where loginname='"& session("adminName") &"')"
			conn.exec sql,"exe"
			session("groupMenu")=session("groupMenu")&", "&rs("Menuid")
		rs.MoveNext
		Loop
	end if
	end if
	pluginInstall
	echo "<"&"script language='javascript'>alert('安装成功');window.parent.frames['topFrame'].location = '../top.asp?'+Math.random();if(window.navigator.userAgent.indexOf(""MSIE"")>=1){top.document.frames.menu.location = '../menu.asp?id=67';}else if(window.navigator.userAgent.indexOf(""Firefox"")>=1){top.document.getElementById('menu').src ='menu.asp?id=67';}else{top.document.getElementById('menu').src ='menu.asp?id=67';}</"&"script>"
	echo "<"&"script> window.location.href='AspCms_plug.asp' </"&"script>"
	end if
End Sub

'插件安装附属方法
Sub pluginInstall
	Dim fso,f,pluginPath,path,s,install
	path = getForm("plugpath","get")
	Set fso=Server.CreateObject(FSO_OBJ_NAME)
	'获取安装文件路径
	pluginPath = Server.MapPath("/plug"&"/"&path)
	if not fso.FolderExists(pluginPath) then exit sub
	if not fso.FileExists(pluginPath&"\#install.inc") then exit sub
	'echo pluginPath
	'读取脚本文件并运行
	Set f = fso.OpenTextFile(pluginPath&"\#install.inc", 1)
	s = f.ReadAll
	f.Close
	Set f = nothing
	Execute(s)	
	'Install server.MapPath("../"),path
	Set fso = Nothing
End Sub

'插件卸载附属方法
Sub pluginUnInstall
	Dim fso,f,pluginPath,path,s,UnInstall
	path = getForm("plugpath","get")	
	Set fso=Server.CreateObject(FSO_OBJ_NAME)
	'获取安装文件路径
	'die  "[" & path & "]"
	pluginPath = Server.MapPath("/plug"&"/"&path)
	if not fso.FolderExists(pluginPath) then exit sub
	if not fso.FileExists(pluginPath&"\#uninstall.inc") then exit sub
	'echo pluginPath
	'读取脚本文件并运行
	Set f = fso.OpenTextFile(pluginPath&"\#uninstall.inc", 1)
	'die f.ReadAll
	s = f.ReadAll
	f.Close
	Set f = nothing
	Execute(s)	
	'UnInstall server.MapPath("../"),path
	Set fso = Nothing
End Sub

Sub plugDel
	dim sql,delid,oldid,newid
	delid=""
	dim plugkey : plugkey=getForm("plugkey","both")
	if isnul(plugkey) then alertMsgAndGo "卸载出错","-1"
	if session("groupMenu")<>"all" then
	Dim rs :set rs =Conn.Exec ("select GroupMenu from {prefix}UserGroup where Groupid in (select Groupid from {prefix}user where loginname='"& session("adminName") &"')","r1")
		
		IF rs.eof or rs.bof Then
			alertMsgAndGo"卸载出错","-1"	
		else
			Do While not rs.eof 
				oldid=rs("GroupMenu")	
				rs.MoveNext		
			Loop
		end if		
	
	set rs =Conn.Exec ("select Menuid from {prefix}Menu where menukey='"& plugkey &"'","r1")
	IF rs.eof or rs.bof Then
		alertMsgAndGo"卸载出错","-1"
	else	
		Do While not rs.eof 
		 	oldid=replace(oldid,", "&rs("Menuid"),"")
			rs.MoveNext
		Loop
	end if
		  
		
	sql="update {prefix}UserGroup set GroupMenu='"&oldid&"' where Groupid in (select Groupid from {prefix}user where loginname='"& session("adminName") &"')"
	conn.exec sql,"exe"
	
	session("groupMenu")=oldid
	end if	
	pluginUnInstall
	conn.exec "delete from {prefix}menu where MenuKey ='"&plugkey&"'","exe"	
	echo "<"&"script language='javascript'>alert('卸载成功');window.parent.frames['topFrame'].location = '../top.asp?'+Math.random();if(window.navigator.userAgent.indexOf(""MSIE"")>=1){top.document.frames.menu.location = '../menu.asp?id=67';}else if(window.navigator.userAgent.indexOf(""Firefox"")>=1){top.document.getElementById('menu').src ='menu.asp?id=67';}else{top.document.getElementById('menu').src ='menu.asp?id=67';}</"&"script>"
	echo "<"&"script> window.location.href='AspCms_plug.asp' </"&"script>"
End Sub	


%>