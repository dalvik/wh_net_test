<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim action : action=getForm("action","get")
if action="add" then addGbook

dim FaqID,FaqTitle,Contact,ContactWay,Content,Reply,AddTime,ReplyTime,FaqStatus,AuditStatus,tab,sql
Sub addGbook
	if getForm("code","post")<>Session("Code") then alertMsgAndGo "验证码不正确","-1"
	'if session("faq") then alertMsgAndGo "请不要重复提交","-1"
	FaqTitle=encodeHtml(filterPara(getForm("FaqTitle","post")))
	Contact=encodeHtml(filterPara(getForm("Contact","post")))
	ContactWay=encodeHtml(filterPara(getForm("ContactWay","post")))
	Content=encodeHtml(filterPara(getForm("Content","post")))
	AddTime=now()
	FaqStatus=false
	AuditStatus=false
	
	tab=split(getForm("tab","post"),",")	
	dim tabStr	: tabStr=""
	dim tabValue  : tabValue=""
	dim valueStr	: valueStr=""
	
	
		dim i	:i=0
		dim rsObj
		sql = "select tabField,tabControlType from {prefix}tabSet order by tabOrder,tabID"
		Set rsObj=Conn.Exec(sql,"r1")
		'0 文本,1 数字,2 编辑器,3 附件,4 日期,5 颜色,6 单选,7 多选
		Do While not rsObj.Eof 		
			tabStr = tabStr&","&rsObj(0)
			if rsObj(1) = 2 then
			valueStr = valueStr & ",'" &encode(getForm(rsObj(0),"post")) & "'"
			else
			valueStr = valueStr & ",'" &getForm(rsObj(0),"post") & "'"
			end if			
			i=i+1
			rsObj.MoveNext
		Loop
		rsObj.Close	:	set rsObj=Nothing
	
	
	
	
	if isnul(FaqTitle) then alertMsgAndGo "问题不能为空","-1"
	if isnul(Contact) then alertMsgAndGo "内容不能为空","-1"
	if isnul(Content) then alertMsgAndGo "联系人不能为空","-1"
	if isnul(ContactWay) then alertMsgAndGo "联系方式不能为空","-1"
	
	
	Conn.Exec"insert into {prefix}GuestBook(LanguageID,FaqTitle,Contact,ContactWay,Content,AddTime,FaqStatus,AuditStatus"&tabStr&") values("&setting.languageID&",'"&FaqTitle&"','"&Contact&"','"&ContactWay&"','"&Content&"','"&AddTime&"',"&FaqStatus&","&AuditStatus&""&valueStr&")","exe"
	session("faq")=true
	
	if messageReminded then sendMail messageAlertsEmail,setting.sitetitle,setting.siteTitle&setting.siteUrl&"--留言信息提醒邮件！","您的网站<a href=""http://"&setting.siteUrl&""">"&setting.siteTitle&"</a>有新的留言信息！<br>"&"<br>问题："&FaqTitle&"<br>内容："&Content&"<br>联系人："&Contact&"<br>联系方式："&ContactWay&"<br>时间："&AddTime	
	
	if SwitchCommentsStatus=0 then
		alertMsgAndGo "留言成功！",getRefer	
	else	
		alertMsgAndGo "留言成功，请等待审核中！",getRefer	
	end if
End Sub

%>