<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
if session("userID")="" then
alertMsgAndGo "����δ��½�����¼��","../member/login.asp"
end if
dim action,needCheck
action=getForm("act","both")
needCheck = false

Select Case LCase(action)	
	Case "buy"
		addOrder()
	Case "comfirm"
		OrderComfirm()
	Case "complete"
		OrderComplete()
	Case "echo"
		echoContent()
	case "directbuy"
		DirectBuy()
	case "print"
		OrderPrint()
	Case Else
		echoContent()
End Select 





Sub OrderComplete		
dim id, sortID,SortAndID
	dim qs
	qs = Session("Cart_QS")
	SortAndID=split(replaceStr(qs,FileExt,""),"_")
	
	if isNul(replaceStr(qs,FileExt,"")) then  echoMsgAndGo "ҳ�治����",3 
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "ҳ�治����",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "ҳ�治����",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "��ѡ���Ʒ��","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "��ѡ���Ʒ��","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productcomfirm.html"
	if not CheckTemplateFile(templatePath) then echo "productcomfirm.html"&err_16
	
	set rsObj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then echoMsgAndGo "ҳ�治����",3 : exit sub		
	with templateObj 
	
		.content=loadFile(templatePath)	
		.parseHtml()
		templateObj.content=replace(templateObj.content,"{aspcms:sortname}",rsObj("SortName"))
		templateObj.content=replace(templateObj.content,"{aspcms:parentsortid}",rsObj("parentid"))		
		templateObj.content=replace(templateObj.content,"{aspcms:sortid}",sortID)	
		templateObj.content=replace(templateObj.content,"{aspcms:topsortid}",rsObj("topsortid"))
		
		if isnul(rsObj("PageKeywords")) then 
			templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",setting.siteKeyWords)	
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",rsObj("PageKeywords"))	
		end if
		if isnul(rsObj("PageDesc")) then
			templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",setting.sitedesc)
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",rsObj("PageDesc"))
		end if
		
		if isnul(rsObj("PageTitle")) then 
			templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("SortName"))	
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("PageTitle"))
		end if
		
		templateObj.parsePosition(sortID)			
		rsObj.close()
		
dim hidden


'������������ {aspcms:order.$trigger$}
templateObj.content=replace(templateObj.content,"{aspcms:order.$trigger$}",hidden)

'�������룺{aspcms:order.no} 
templateObj.content=replace(templateObj.content,"{aspcms:order.no}",Session("Print.OrderNo"))
'�������ڣ�{aspcms:order.date} 
templateObj.content=replace(templateObj.content,"{aspcms:order.date}",Session("Print.OrderDate"))
'�� �� ����{aspcms:order.username} 
templateObj.content=replace(templateObj.content,"{aspcms:order.username}",Session("Print.OrderUserName"))
'����״̬��{aspcms:order.state} 
templateObj.content=replace(templateObj.content,"{aspcms:order.state}","δ����(�ɹ��µ�)")
'�����ܽ�{aspcms:order.total}Ԫ 
templateObj.content=replace(templateObj.content,"{aspcms:order.total}",Session("Print.OrderTotal"))
'��Ʒ��������{aspcms:order.count} 	
templateObj.content=replace(templateObj.content,"{aspcms:order.count}",Session("Print.OrderCount"))
'�ͻ�������{aspcms:order.nicename}
templateObj.content=replace(templateObj.content,"{aspcms:order.nicename}",Session("Print.OrderNiceName"))
'��ϵ�绰��{aspcms:order.tel}
templateObj.content=replace(templateObj.content,"{aspcms:order.tel}",Session("Print.OrderTel"))
'�ֻ����룺{aspcms:order.cellphone}
templateObj.content=replace(templateObj.content,"{aspcms:order.cellphone}",Session("Print.OrderCellphone"))
'�������룺{aspcms:order.zipcode}
templateObj.content=replace(templateObj.content,"{aspcms:order.zipcode}",Session("Print.OrderZipcode"))
'��ϵ��ַ��{aspcms:order.address}
templateObj.content=replace(templateObj.content,"{aspcms:order.address}",Session("Print.OrderAddress"))
'�������䣺{aspcms:order.email}
templateObj.content=replace(templateObj.content,"{aspcms:order.email}",Session("Print.OrderEmail"))
'�����ԣ�{aspcms:order.note}
templateObj.content=replace(templateObj.content,"{aspcms:order.note}",Session("Print.OrderNote"))
'֧����ʽ
templateObj.content=replace(templateObj.content,"{aspcms:order.payment}",Session("Print.OrderPayMent"))

select case Session("Print.OrderPayMent")
	case "����֧��"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Online)
	case "����֧��"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Bank)
	case "�ʾָ���"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_PostOffice)
end select


'��Ʒ�б�
templateObj.content=replace(templateObj.content,"{aspcms:selectproduct}",SelectProductList)
 

		.parseCommon() 		
		echo .content 
	end with
	
	set templateobj =nothing : terminateAllObjects
	
End Sub

'=======================================
'ֱ�ӹ���
'=======================================
Sub DirectBuy
Dim dicCount,dicPrice
Dim proid,count,proprice
Session.Contents.Remove("Cart")
Session.Contents.Remove("Cart.Price")

proid = filterPara(getForm("proid","post"))
count = getForm("count","post")
proprice = getForm("proprice","post")


Set dicCount = Server.CreateObject("Scripting.Dictionary")
Set dicPrice = Server.CreateObject("Scripting.Dictionary")
dicCount.add proid,count
dicPrice.add proid,proprice
Set Session("Cart") = dicCount
Set Session("Cart.Price") = dicPrice

Session("Cart_QS") = Request.QueryString
echoContent()	
End Sub


'=======================================
'������Ϣ��д
'=======================================
Sub echoContent()
	'if rCookie("loginstatus")<>"1" then
	'	response.Redirect "../member/login.asp"
	'end if
	dim total
	dim id, sortID,SortAndID
	dim qs	
	qs = Session("Cart_QS")	
	SortAndID=split(replaceStr(qs,FileExt,""),"_")
	if isNul(replaceStr(qs,FileExt,"")) then  echoMsgAndGo "ҳ�治����",3 
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "ҳ�治����",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "ҳ�治����",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "��ѡ���Ʒ��","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "��ѡ���Ʒ��","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productcheckout.html"
	if not CheckTemplateFile(templatePath) then echo "productcheckout.html"&err_16
	
	set rsobj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then echoMsgAndGo "ҳ�治����",3 : exit sub		
	with templateObj 
		.content=loadFile(templatePath)	
		.parseHtml()
		templateObj.content=replace(templateObj.content,"{aspcms:sortname}",rsObj("SortName"))
		templateObj.content=replace(templateObj.content,"{aspcms:parentsortid}",rsObj("parentid"))		
		templateObj.content=replace(templateObj.content,"{aspcms:sortid}",sortID)	
		templateObj.content=replace(templateObj.content,"{aspcms:topsortid}",rsObj("topsortid"))
		if isnul(rsObj("PageKeywords")) then 
			templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",setting.siteKeyWords)	
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",rsObj("PageKeywords"))	
		end if
		if isnul(rsObj("PageDesc")) then
			templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",setting.sitedesc)
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",rsObj("PageDesc"))
		end if
		if isnul(rsObj("PageTitle")) then 
			templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("SortName"))	
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("PageTitle"))
		end if
		templateObj.parsePosition(sortID)
		
	
	
	Dim m_username,gender
	if isnul(session("loginstatus")) then  session("loginstatus")=0
	if session("loginstatus")="1" then  
		set rsObj=conn.Exec("select * from {prefix}User where UserID="&trim(session("userID")),"r1")
		m_username = rsObj("loginname")
	else 
		gender=1
	end if
	rsObj.close()
		
		.content=replaceStr(.content,"{aspcms:selectproduct}",SelectProductList)
		randomize 		
		.content=replaceStr(.content,"{aspcms:order.orderno}",Year(Now)&MOnth(Now)&Day(Now)&Int(10000000*   Rnd))
		.content=replaceStr(.content,"{aspcms:order.username}",m_username)
		.content=replaceStr(.content,"{aspcms:order.total}",session("total"))
		.parseCommon() 
		echo .content 
	end with
	set templateobj =nothing : terminateAllObjects
	'echo "ajaaj"
End Sub


'=======================================
'��Ӷ���
'=======================================
Sub addOrder
if needCheck then
if getForm("code","post")<>Session("Code") then alertMsgAndGo "��֤�벻��ȷ","-1"
end if


dim orderno, username,  m_state, payment, nicename, tel,cellphone, zipcode, address, email, note, m_to, invoice,addtime
dim sql

	orderno=getForm("orderno","post")
	username=filterPara(getForm("username","post"))	
	nicename=filterPara(getForm("nicename","post"))	
	m_state=0
	address=filterPara(getForm("address","post"))	
	zipcode=filterPara(getForm("zipcode","post"))
	cellphone=filterPara(getForm("cellphone","post"))
	tel=filterPara(getForm("tel","post"))	
	email=filterPara(getForm("email","post"))
	m_to=filterPara(getForm("to","post"))	
	note=filterPara(getForm("note","post"))	
	invoice=filterPara(getForm("invoice","post"))
	payment=filterPara(getForm("payment","post"))
	 
	addtime = now
	
	if isnul(nicename) then alertMsgAndGo "��ϵ�˲���Ϊ��","-1"
		if isnul(tel) then alertMsgAndGo "��ϵ�绰����Ϊ��","-1"
		if isnul(cellphone) then alertMsgAndGo "�ֻ�����Ϊ��","-1"
		
		if username = "" then username = "�ο�"
	dim userid:userid=session("userID")
	if session("norefresh") <> orderno then
	if isnul(userid) then userid=0
		session("norefresh") = orderno
	sql = "INSERT INTO {prefix}order2 ( orderno, username, ordertime, state, payment, nicename, tel, cellphone, zipcode, address, email, [note], [to], invoice,userid)VALUES ('"&orderno&"', '"&username&"', now(), "&m_state&", '"&payment&"', '"&nicename&"', '"&tel&"', '"&cellphone&"', '"&zipcode&"', '"&address&"', '"&email&"', '"&note&"', '"&m_to&"', "&invoice&", "&userid&")"
	
	Conn.Exec sql,"exe"		
	dim dic,md,dicPrice
	Set dic = Session("Cart")
	Set dicPrice=Session("Cart.Price")
	for each md in dic
	dim prices
		if isnul(dicPrice(md)) then  prices=0 else prices=dicPrice(md)
		sql = "insert into {prefix}OrderProduct (orderno,productid,[count],instantprice) values ('"&orderno&"',"&md&","&dic(md)&","&prices&")"
		conn.exec sql,"exe"
		'echo sql & "<br>"
	next	
	'Session.Contents.Remove("Cart")
	'Session.Contents.Remove("Cart.Price")
else
 
	
	end if
	
	
	
	
	
	if orderReminded then sendMail messageAlertsEmail,setting.sitetitle,setting.siteTitle&setting.siteUrl&"--������Ϣ�����ʼ���","������վ<a href=""http://"&setting.siteUrl&""">"&setting.siteTitle&"</a>���µĶ�����Ϣ��<br>������ţ�"&orderno&"<br>��Ա�ʺţ�"&username&"<br>��ϵ�ˣ�"&nicename&"<br>�绰���룺"&tel&"<br>�ֻ����룺"&cellphone&"<br>�������䣺"&Email&"<br>��ע��"&note&"<br>����ʱ�䣺"&addtime
	
	alertMsgAndGo "��Ʒ�����ɹ���","?act=complete"
	if Session("Cart_QS") <> "" then Session.Contents.Remove("Cart_QS")
	Session.Contents.Remove("norefresh")
End Sub



'=======================================
'������ϸ�б�
'=======================================
Function SelectProductList
Dim result,sql,rs,md,dic
Dim total:total=0
result=""
result = result & "<table width='100%' cellspacing='0' cellpadding='0' style='background:none; border:1px solid #CCCCCC;'>"
result = result & "<tr style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;' align='center'>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'>��ƷͼƬ</td>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'>��Ʒ����</td>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'>����</td>"
result = result & "<td style='background:none; border-bottom:1px solid #CCCCCC;'>����</td>"
result = result & "</tr>"

Set dic = Session("Cart")
	For Each md In dic
		sql = "select P_Price,IndexImage,title,contentid from {prefix}Content where contentid=" & md
		Set rs = GetRS(sql)	
result = result & "<tr style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;' align='center'>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'><a title='����鿴ԭͼ' href='../content/?"&rs("contentid")&".html'  target=_blank><img src='"&rs("IndexImage")&"' width='70' height='50'style='border:none' /></a></td>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'><a title='"&rs("title")&"' href='../content/?"&rs("contentid")&".html'  target=_blank>"&rs("title")&"</a></td>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'><span id='price"&md&"'>"&rs("P_Price")&"</span>Ԫ</td>"
result = result & "<td style='background:none; border-bottom:1px solid #CCCCCC;'>"&dic(md)&"</td>"
result = result & "</tr>"
	total = total + rs("P_Price") * dic(md)
		rs.close:set rs = nothing
		session("total")=total
	Next
result = result & "<tr bgcolor='#FFFFFF' align='center'><td colspan='4' bgcolor='#FFFFFF' align='center'>�ܼ۸�"&total&"Ԫ</td></tr>"

result = result & "</table>"
SelectProductList = result
End Function
'=======================================
'��Ӷ���
'=======================================
Function GetRS(sql)

On Error Resume Next
Set GetRS = conn.exec(sql,"r1")
If Err Then
	'echo sql
	'die err.description
	die "���󣬱�ṹ�������°汾"
End If
End Function



'=======================================
'������ӡ
'=======================================
Sub OrderPrint
if needCheck then
if getForm("code","post")<>Session("Code") then alertMsgAndGo "��֤�벻��ȷ","-1"
end if
dim sql		
dim id, sortID,SortAndID
	dim qs
	qs = Session("Cart_QS")
	SortAndID=split(replaceStr(qs,FileExt,""),"_")
	
	if isNul(replaceStr(qs,FileExt,"")) then  echoMsgAndGo "ҳ�治����",3 
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "ҳ�治����",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "ҳ�治����",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "��ѡ���Ʒ��","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "��ѡ���Ʒ��","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productprint.html"
	if not CheckTemplateFile(templatePath) then echo "productprint.html"&err_16
	
	set rsObj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then echoMsgAndGo "ҳ�治����",3 : exit sub		
	with templateObj 
	
		.content=loadFile(templatePath)	
		.parseHtml()
		


'�������룺{aspcms:order.no} 
templateObj.content=replace(templateObj.content,"{aspcms:order.no}",Session("Print.OrderNo"))
'�������ڣ�{aspcms:order.date} 
templateObj.content=replace(templateObj.content,"{aspcms:order.date}",Session("Print.OrderDate"))
'�� �� ����{aspcms:order.username} 
templateObj.content=replace(templateObj.content,"{aspcms:order.username}",Session("Print.OrderUserName"))
'����״̬��{aspcms:order.state} 
templateObj.content=replace(templateObj.content,"{aspcms:order.state}",Session("Print.OrderState"))
'�����ܽ�{aspcms:order.total}Ԫ 
templateObj.content=replace(templateObj.content,"{aspcms:order.total}",Session("Print.OrderTotal"))
'��Ʒ��������{aspcms:order.count} 	
templateObj.content=replace(templateObj.content,"{aspcms:order.count}",Session("Print.OrderCount"))
'�ͻ�������{aspcms:order.nicename}
templateObj.content=replace(templateObj.content,"{aspcms:order.nicename}",Session("Print.OrderNiceName"))
'��ϵ�绰��{aspcms:order.tel}
templateObj.content=replace(templateObj.content,"{aspcms:order.tel}",Session("Print.OrderTel"))
'�ֻ����룺{aspcms:order.cellphone}
templateObj.content=replace(templateObj.content,"{aspcms:order.cellphone}",Session("Print.OrderCellphone"))
'�������룺{aspcms:order.zipcode}
templateObj.content=replace(templateObj.content,"{aspcms:order.zipcode}",Session("Print.OrderZipcode"))
'��ϵ��ַ��{aspcms:order.address}
templateObj.content=replace(templateObj.content,"{aspcms:order.address}",Session("Print.OrderAddress"))
'�������䣺{aspcms:order.email}
templateObj.content=replace(templateObj.content,"{aspcms:order.email}",Session("Print.OrderEmail"))
'�����ԣ�{aspcms:order.note}
templateObj.content=replace(templateObj.content,"{aspcms:order.note}",Session("Print.OrderNote"))
'֧����ʽ
templateObj.content=replace(templateObj.content,"{aspcms:order.payment}",Session("Print.OrderPayMent"))

select case Session("Print.OrderPayMent")
	case "����֧��"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Online)
	case "����֧��"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Bank)
	case "�ʾָ���"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_PostOffice)
end select


'��Ʒ�б�
templateObj.content=replace(templateObj.content,"{aspcms:selectproduct}",SelectProductList)
 

		.parseCommon() 		
		echo .content 
	end with
	
	set templateobj =nothing : terminateAllObjects
End Sub



'=======================================
'����ȷ��
'=======================================
Sub OrderComfirm()
if needCheck then
if getForm("code","post")<>Session("Code") then alertMsgAndGo "��֤�벻��ȷ","-1"
end if


dim orderno, username,  m_state, payment, nicename, tel,cellphone, zipcode, address, email, note, m_to, invoice,addtime 
dim sql

	orderno=getForm("orderno","post")
	username=filterPara(getForm("username","post"))	
	nicename=filterPara(getForm("nicename","post"))	
	m_state=0
	address=filterPara(getForm("address","post"))	
	zipcode=filterPara(getForm("zipcode","post"))
	cellphone=filterPara(getForm("cellphone","post"))
	tel=filterPara(getForm("tel","post"))	
	email=filterPara(getForm("email","post"))
	m_to=filterPara(getForm("to","post"))	
	note=filterPara(getForm("note","post"))	
	invoice=filterPara(getForm("invoice","post"))
	Payment=filterPara(getForm("Payment","post"))
	 
	addtime = now
	
if isnul(nicename) then alertMsgAndGo "��ϵ�˲���Ϊ��","-1"
if isnul(tel) then alertMsgAndGo "��ϵ�绰����Ϊ��","-1"
if isnul(cellphone) then alertMsgAndGo "�ֻ�����Ϊ��","-1"
		
		
dim id, sortID,SortAndID
	dim qs
	qs = Session("Cart_QS")
	SortAndID=split(replaceStr(qs,FileExt,""),"_")
	
	if isNul(replaceStr(qs,FileExt,"")) then  echoMsgAndGo "ҳ�治����",3 
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "ҳ�治����",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "ҳ�治����",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "��ѡ���Ʒ��","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "��ѡ���Ʒ��","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productcomfirm.html"
	if not CheckTemplateFile(templatePath) then echo "productcomfirm.html"&err_16
	
	set rsObj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then echoMsgAndGo "ҳ�治����",3 : exit sub		
	with templateObj 
	
		.content=loadFile(templatePath)	
		.parseHtml()
		templateObj.content=replace(templateObj.content,"{aspcms:sortname}",rsObj("SortName"))
		templateObj.content=replace(templateObj.content,"{aspcms:parentsortid}",rsObj("parentid"))		
		templateObj.content=replace(templateObj.content,"{aspcms:sortid}",sortID)	
		templateObj.content=replace(templateObj.content,"{aspcms:topsortid}",rsObj("topsortid"))
		
		if isnul(rsObj("PageKeywords")) then 
			templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",setting.siteKeyWords)	
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",rsObj("PageKeywords"))	
		end if
		if isnul(rsObj("PageDesc")) then
			templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",setting.sitedesc)
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",rsObj("PageDesc"))
		end if
		
		if isnul(rsObj("PageTitle")) then 
			templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("SortName"))	
		else
			templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("PageTitle"))
		end if
		
		templateObj.parsePosition(sortID)	
		
		Dim m_username,gender
		if isnul(session("loginstatus")) then  session("loginstatus")=0
		if session("loginstatus")="1" then  
			set rsObj=conn.Exec("select * from {prefix}User where UserID="&trim(session("userID")),"r1")
			m_username = rsObj("loginname")
		else 
			gender=1
		end if
		rsObj.close()
		
		
Dim dic,dicPrice,d,oTotal,oCount
oTotal = 0:oCount = 0
Set dic = Session("Cart")
Set dicPrice = Session("Cart.Price")
For Each d In dic
	dim prices
	if isnul(dicPrice(d)) then  prices=0 else prices=dicPrice(d)
	oTotal = oTotal + Eval(dic(d) * prices)
	oCount = oCount + dic(d)
Next

Set dic = nothing:Set dicPrice = nothing


dim hidden
hidden = "<form action='?act=buy' method='post' name='frmComfirm'>"
hidden = hidden & "<input type='hidden' name='orderno' value='"&orderno&"' />"
hidden = hidden & "<input type='hidden' name='username' value='"&username&"' />"
hidden = hidden & "<input type='hidden' name='nicename' value='"&nicename&"' />"
hidden = hidden & "<input type='hidden' name='address' value='"&address&"' />"
hidden = hidden & "<input type='hidden' name='zipcode' value='"&zipcode&"' />"
hidden = hidden & "<input type='hidden' name='cellphone' value='"&cellphone&"' />"
hidden = hidden & "<input type='hidden' name='tel' value='"&tel&"' />"
hidden = hidden & "<input type='hidden' name='email' value='"&email&"' />"
hidden = hidden & "<input type='hidden' name='to' value='"&m_to&"' />"
hidden = hidden & "<input type='hidden' name='note' value='"&note&"' />"
hidden = hidden & "<input type='hidden' name='invoice' value='"&invoice&"' />"
hidden = hidden & "<input type='hidden' name='Payment' value='"&Payment&"' />"
'hidden = hidden & "<input type='hidden' name='state' value='0' />"
hidden = hidden & "<p align='center'>"
hidden = hidden & "<input type='submit' value='ȷ�϶���' name='btnSubmit'/></p></form>"

'������������ {aspcms:order.$trigger$}
templateObj.content=replace(templateObj.content,"{aspcms:order.$trigger$}",hidden)

If m_username = "" then m_username = "�ο�"

'�������룺{aspcms:order.no} 
templateObj.content=replace(templateObj.content,"{aspcms:order.no}",orderno)
Session("Print.OrderNo") = orderno
'�������ڣ�{aspcms:order.date} 
templateObj.content=replace(templateObj.content,"{aspcms:order.date}",addtime)
Session("Print.OrderDate") = addtime
'�� �� ����{aspcms:order.username} 
templateObj.content=replace(templateObj.content,"{aspcms:order.username}",m_username)
Session("Print.OrderUserName") = m_username
'����״̬��{aspcms:order.state} 
templateObj.content=replace(templateObj.content,"{aspcms:order.state}","δ����")
Session("Print.OrderState") = "δ����"
'�����ܽ�{aspcms:order.total}Ԫ 
templateObj.content=replace(templateObj.content,"{aspcms:order.total}",oTotal)
Session("Print.OrderTotal") = oTotal
'��Ʒ��������{aspcms:order.count} 	
templateObj.content=replace(templateObj.content,"{aspcms:order.count}",oCount)
Session("Print.OrderCount") = oCount
'�ͻ�������{aspcms:order.nicename}
templateObj.content=replace(templateObj.content,"{aspcms:order.nicename}",nicename)
Session("Print.OrderNiceName") = nicename
'��ϵ�绰��{aspcms:order.tel}
templateObj.content=replace(templateObj.content,"{aspcms:order.tel}",tel)
Session("Print.OrderTel") = tel
'�ֻ����룺{aspcms:order.cellphone}
templateObj.content=replace(templateObj.content,"{aspcms:order.cellphone}",cellphone)
Session("Print.OrderCellphone") = cellphone
'�������룺{aspcms:order.zipcode}
templateObj.content=replace(templateObj.content,"{aspcms:order.zipcode}",zipcode)
Session("Print.OrderZipcode") = zipcode
'��ϵ��ַ��{aspcms:order.address}
templateObj.content=replace(templateObj.content,"{aspcms:order.address}",address)
Session("Print.OrderAddress") = address
'�������䣺{aspcms:order.email}
templateObj.content=replace(templateObj.content,"{aspcms:order.email}",email)
Session("Print.OrderEmail") = email
'�����ԣ�{aspcms:order.note}
templateObj.content=replace(templateObj.content,"{aspcms:order.note}",note)
Session("Print.OrderNote") = note
'֧����ʽ
templateObj.content=replace(templateObj.content,"{aspcms:order.payment}",Payment)
Session("Print.OrderPayMent") = Payment

select case Payment
	case "����֧��"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Online)
	case "����֧��"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Bank)
	case "�ʾָ���"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_PostOffice)
end select


'��Ʒ�б�
templateObj.content=replace(templateObj.content,"{aspcms:selectproduct}",SelectProductList)
 

		.parseCommon() 		
		echo .content 
	end with
	
	set templateobj =nothing : terminateAllObjects
	
End Sub
%>
