<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim userID,Page,qs,Password,rePassword,PswQuestion,PswAnswer,TrueName,Gender,Birthday,Country,Province,City,Address,PostCode,Phone,Mobile,Email,QQ,MSN,sql
userID=session("userID")
if isNul(userID) or not isNum(userID) then alertMsgAndGo "�Բ���,���ĵ�¼״̬�Ѿ�ʧЧ,�����µ�¼!","login.asp"


dim action : action=getform("action","get")

if action = "update" then
	Password=getForm("Password","post")
	rePassword=getForm("rePassword","post")
	PswQuestion=getForm("PswQuestion","post")
	PswAnswer=getForm("PswAnswer","post")
	TrueName=getForm("TrueName","post")
	Gender=getForm("Gender","post")
	Birthday=getForm("Birthday","post")
	Country=getForm("Country","post")
	Province=getForm("Province","post")
	City=getForm("City","post")
	Address=getForm("Address","post")
	PostCode=getForm("PostCode","post")
	Phone=getForm("Phone","post")
	Mobile=getForm("Mobile","post")
	Email=getForm("Email","post")
	QQ=getForm("QQ","post")
	MSN=getForm("MSN","post")	
	
	if not isNum(Gender) then alertMsgAndGo "�Ա��ʶ����ȷ��","-1"	
	if isnul(Birthday) then Birthday="1900-01-01"	
	Dim passStr
	if isNul(Password) then 
		passStr=""
	else
		if Password=rePassword then
		passStr=" , [Password]='"&md5(Password, 16)&"'"
		else
		alertMsgAndGo "�����������벻һ�£�","-1"
		end if
		
	end if
	
	if not isnum(Gender) then Gender = 1
	if not isnum(Birthday) then Birthday = 1
	
	sql="update {prefix}User set PswQuestion='"&PswQuestion&"',PswAnswer='"&PswAnswer&"',TrueName='"&TrueName&"',Gender="&Gender&",Birthday="&Birthday&",Country='"&Country&"',Province='"&Province&"',City='"&City&"',Address='"&Address&"',PostCode='"&PostCode&"',Phone='"&Phone&"',Mobile='"&Mobile&"',Email='"&Email&"',QQ='"&QQ&"',MSN='"&MSN&"'"&passStr&" where userid="&userID
	'die sql

	conn.exec sql,"exe"	
	alertMsgAndGo "�޸ĳɹ�","../index.asp"
else	
	dim content : content=makeuser(userID, Page, false)
	die content
	if isnul(content) then 
		echoMsgAndGo "ҳ�治����",3 
	else
		echo content
	end If


end if
If DebugMode Then echo timer - AppSpan
%>