<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim action : action=getForm("action","get")
if action="add" then addform

dim formID,formTitle,Contact,ContactWay,Content,Reply,AddTime,ReplyTime,formStatus,AuditStatus,aform,sql
Sub addform
	if getForm("code","post")<>Session("Code") then alertMsgAndGo "验证码不正确","-1"
	'if session("aform") then alertMsgAndGo "请不要重复提交","-1"
	formTitle=filterPara(getForm("formTitle","post"))
	Contact=filterPara(getForm("Contact","post"))
	ContactWay=filterPara(getForm("ContactWay","post"))
	Content=encode(filterPara(getForm("Content","post")))
	AddTime=now()
	formStatus=false
	AuditStatus=false
	
	aform=split(getForm("aform","post"),",")	
	dim formStr	: formStr=""
	dim formValue  : formValue=""
	dim valueStr	: valueStr=""
	
	
		dim i	:i=0
		dim rsObj
		sql = "select formField,formControlType from {prefix}formSet order by formOrder,formID"
		Set rsObj=Conn.Exec(sql,"r1")
		'0 文本,1 数字,2 编辑器,3 附件,4 日期,5 颜色,6 单选,7 多选
		Do While not rsObj.Eof 		
			formStr = formStr&","&rsObj(0)
			if rsObj(1) = 2 then
			valueStr = valueStr & ",'" &encode(getForm(rsObj(0),"post")) & "'"
			else
			valueStr = valueStr & ",'" &getForm(rsObj(0),"post") & "'"
			end if			
			i=i+1
			rsObj.MoveNext
		Loop
		rsObj.Close	:	set rsObj=Nothing
		
	
	Conn.Exec"insert into {prefix}custom(LanguageID,AddTime,customStatus,AuditStatus"&formStr&") values("&setting.languageID&",'"&AddTime&"',"&formStatus&","&AuditStatus&""&valueStr&")","exe"
	'session("aform")=true
	
	
		alertMsgAndGo "提交成功！",getRefer	
	
End Sub

%>