<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim action : action=getForm("action","get")
if action="add" then addGbook

dim FaqID,FaqTitle,Contact,ContactWay,Content,Reply,AddTime,ReplyTime,FaqStatus,AuditStatus,tab,sql
Sub addGbook
	if getForm("code","post")<>Session("Code") then alertMsgAndGo "��֤�벻��ȷ","-1"
	'if session("faq") then alertMsgAndGo "�벻Ҫ�ظ��ύ","-1"
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
		'0 �ı�,1 ����,2 �༭��,3 ����,4 ����,5 ��ɫ,6 ��ѡ,7 ��ѡ
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
	
	
	
	
	if isnul(FaqTitle) then alertMsgAndGo "���ⲻ��Ϊ��","-1"
	if isnul(Contact) then alertMsgAndGo "���ݲ���Ϊ��","-1"
	if isnul(Content) then alertMsgAndGo "��ϵ�˲���Ϊ��","-1"
	if isnul(ContactWay) then alertMsgAndGo "��ϵ��ʽ����Ϊ��","-1"
	
	
	Conn.Exec"insert into {prefix}GuestBook(LanguageID,FaqTitle,Contact,ContactWay,Content,AddTime,FaqStatus,AuditStatus"&tabStr&") values("&setting.languageID&",'"&FaqTitle&"','"&Contact&"','"&ContactWay&"','"&Content&"','"&AddTime&"',"&FaqStatus&","&AuditStatus&""&valueStr&")","exe"
	session("faq")=true
	
	if messageReminded then sendMail messageAlertsEmail,setting.sitetitle,setting.siteTitle&setting.siteUrl&"--������Ϣ�����ʼ���","������վ<a href=""http://"&setting.siteUrl&""">"&setting.siteTitle&"</a>���µ�������Ϣ��<br>"&"<br>���⣺"&FaqTitle&"<br>���ݣ�"&Content&"<br>��ϵ�ˣ�"&Contact&"<br>��ϵ��ʽ��"&ContactWay&"<br>ʱ�䣺"&AddTime	
	
	if SwitchCommentsStatus=0 then
		alertMsgAndGo "���Գɹ���",getRefer	
	else	
		alertMsgAndGo "���Գɹ�����ȴ�����У�",getRefer	
	end if
End Sub

%>