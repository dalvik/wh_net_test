<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim action : action=getform("action","get")
if action = "login" then
	dim LoginName,Password,sql,code,rs
	
	if getForm("code","post")<>Session("Code") then alertMsgAndGo "��֤�벻��ȷ","-1"
	LoginName = filterPara(getForm("LoginName","post"))
	Password = md5(getForm("userPass","post"),16)
	
	sql = "select count(*) from {prefix}User where LoginName = '"& LoginName &"' and PassWord='"&Password&"'"
	Dim rsObj : Set rsObj=Conn.Exec(sql,"r1")
	if rsObj(0)=1 then
		Set rsObj=Conn.Exec("select UserId,LoginName,GroupMark,{prefix}User.GroupID as ugi from {prefix}User,{prefix}UserGroup where {prefix}User.GroupID={prefix}UserGroup.GroupID and LoginName='"&LoginName&"' and UserStatus=1","r1")	
		if not rsObj.Eof Then
		'die rsObj("LoginName")	
			session("loginName")=rsObj("LoginName")	
			session("userID")=rsObj("UserId")	
			session("GroupID")=rsObj("ugi")	
			session("loginstatus")="1"
			session("GroupMark")=rsObj("GroupMark")
			
			Conn.Exec"update {prefix}User set LastLoginTime='"&now()&"',LastLoginIP='"&getIp()&"',LoginCount=LoginCount+1 where UserId="&rsObj("UserId"),"exe"			
			'Response.Redirect Request.ServerVariables("HTTP_REFERER")
			response.Redirect("userinfo.asp")
			response.End()
		else
			alertMsgAndGo "�Բ��������˺��ѱ�����!","-1"
		end if
	else
		alertMsgAndGo "�û������������,ϵͳ�������ص�¼ҳ��!","login.asp"
	end if
	rsObj.Close() : set rsObj=Nothing	
elseif action = "logout" then
	session("loginName")=""	
	session("loginstatus")="0"
	session("userID")=""
	session("GroupID")=2
	session("GroupMark")=""
	alertMsgAndGo "���Ѿ��ɹ��˳���¼!",sitePath&setting.languagePath
elseif action = "relog" then
	alertMsgAndGo "�Բ���,���ĵ�¼״̬�Ѿ�ʧЧ,�����µ�¼!",sitePath&setting.languagePath
else
	echoContent()
end if

Sub echoContent()
	dim templateobj,templatePath : set templateobj = new TemplateClass	
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/login.html"	
	if not CheckTemplateFile(templatePath) then echo "login.html"&err_16	
	
	with templateObj 
		.content=loadFile(templatePath) 
		.parseHtml()		
		.parseCommon
		echo .content 
	end with	
	set templateobj =nothing : terminateAllObjects
End Sub
%>