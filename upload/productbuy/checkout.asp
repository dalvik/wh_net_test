<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
if session("userID")="" then
alertMsgAndGo "您还未登陆，请登录！","../member/login.asp"
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
	
	if isNul(replaceStr(qs,FileExt,"")) then  echoMsgAndGo "页面不存在",3 
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "页面不存在",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "页面不存在",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "请选择产品！","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "请选择产品！","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productcomfirm.html"
	if not CheckTemplateFile(templatePath) then echo "productcomfirm.html"&err_16
	
	set rsObj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then echoMsgAndGo "页面不存在",3 : exit sub		
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


'订单参数传递 {aspcms:order.$trigger$}
templateObj.content=replace(templateObj.content,"{aspcms:order.$trigger$}",hidden)

'订单号码：{aspcms:order.no} 
templateObj.content=replace(templateObj.content,"{aspcms:order.no}",Session("Print.OrderNo"))
'订单日期：{aspcms:order.date} 
templateObj.content=replace(templateObj.content,"{aspcms:order.date}",Session("Print.OrderDate"))
'用 户 名：{aspcms:order.username} 
templateObj.content=replace(templateObj.content,"{aspcms:order.username}",Session("Print.OrderUserName"))
'订单状态：{aspcms:order.state} 
templateObj.content=replace(templateObj.content,"{aspcms:order.state}","未处理(成功下单)")
'订单总金额：{aspcms:order.total}元 
templateObj.content=replace(templateObj.content,"{aspcms:order.total}",Session("Print.OrderTotal"))
'商品总数量：{aspcms:order.count} 	
templateObj.content=replace(templateObj.content,"{aspcms:order.count}",Session("Print.OrderCount"))
'客户姓名：{aspcms:order.nicename}
templateObj.content=replace(templateObj.content,"{aspcms:order.nicename}",Session("Print.OrderNiceName"))
'联系电话：{aspcms:order.tel}
templateObj.content=replace(templateObj.content,"{aspcms:order.tel}",Session("Print.OrderTel"))
'手机号码：{aspcms:order.cellphone}
templateObj.content=replace(templateObj.content,"{aspcms:order.cellphone}",Session("Print.OrderCellphone"))
'邮政编码：{aspcms:order.zipcode}
templateObj.content=replace(templateObj.content,"{aspcms:order.zipcode}",Session("Print.OrderZipcode"))
'联系地址：{aspcms:order.address}
templateObj.content=replace(templateObj.content,"{aspcms:order.address}",Session("Print.OrderAddress"))
'电子邮箱：{aspcms:order.email}
templateObj.content=replace(templateObj.content,"{aspcms:order.email}",Session("Print.OrderEmail"))
'简单留言：{aspcms:order.note}
templateObj.content=replace(templateObj.content,"{aspcms:order.note}",Session("Print.OrderNote"))
'支付方式
templateObj.content=replace(templateObj.content,"{aspcms:order.payment}",Session("Print.OrderPayMent"))

select case Session("Print.OrderPayMent")
	case "在线支付"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Online)
	case "银行支付"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Bank)
	case "邮局付款"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_PostOffice)
end select


'商品列表
templateObj.content=replace(templateObj.content,"{aspcms:selectproduct}",SelectProductList)
 

		.parseCommon() 		
		echo .content 
	end with
	
	set templateobj =nothing : terminateAllObjects
	
End Sub

'=======================================
'直接购买
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
'订单信息填写
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
	if isNul(replaceStr(qs,FileExt,"")) then  echoMsgAndGo "页面不存在",3 
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "页面不存在",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "页面不存在",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "请选择产品！","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "请选择产品！","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productcheckout.html"
	if not CheckTemplateFile(templatePath) then echo "productcheckout.html"&err_16
	
	set rsobj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then echoMsgAndGo "页面不存在",3 : exit sub		
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
'添加订单
'=======================================
Sub addOrder
if needCheck then
if getForm("code","post")<>Session("Code") then alertMsgAndGo "验证码不正确","-1"
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
	
	if isnul(nicename) then alertMsgAndGo "联系人不能为空","-1"
		if isnul(tel) then alertMsgAndGo "联系电话不能为空","-1"
		if isnul(cellphone) then alertMsgAndGo "手机不能为空","-1"
		
		if username = "" then username = "游客"
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
	
	
	
	
	
	if orderReminded then sendMail messageAlertsEmail,setting.sitetitle,setting.siteTitle&setting.siteUrl&"--订单信息提醒邮件！","您的网站<a href=""http://"&setting.siteUrl&""">"&setting.siteTitle&"</a>有新的订单信息！<br>订单编号："&orderno&"<br>会员帐号："&username&"<br>联系人："&nicename&"<br>电话号码："&tel&"<br>手机号码："&cellphone&"<br>电子信箱："&Email&"<br>备注："&note&"<br>申请时间："&addtime
	
	alertMsgAndGo "产品订购成功！","?act=complete"
	if Session("Cart_QS") <> "" then Session.Contents.Remove("Cart_QS")
	Session.Contents.Remove("norefresh")
End Sub



'=======================================
'订单详细列表
'=======================================
Function SelectProductList
Dim result,sql,rs,md,dic
Dim total:total=0
result=""
result = result & "<table width='100%' cellspacing='0' cellpadding='0' style='background:none; border:1px solid #CCCCCC;'>"
result = result & "<tr style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;' align='center'>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'>商品图片</td>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'>商品名称</td>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'>单价</td>"
result = result & "<td style='background:none; border-bottom:1px solid #CCCCCC;'>数量</td>"
result = result & "</tr>"

Set dic = Session("Cart")
	For Each md In dic
		sql = "select P_Price,IndexImage,title,contentid from {prefix}Content where contentid=" & md
		Set rs = GetRS(sql)	
result = result & "<tr style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;' align='center'>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'><a title='点击查看原图' href='../content/?"&rs("contentid")&".html'  target=_blank><img src='"&rs("IndexImage")&"' width='70' height='50'style='border:none' /></a></td>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'><a title='"&rs("title")&"' href='../content/?"&rs("contentid")&".html'  target=_blank>"&rs("title")&"</a></td>"
result = result & "<td style='background:none; border-right:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;'><span id='price"&md&"'>"&rs("P_Price")&"</span>元</td>"
result = result & "<td style='background:none; border-bottom:1px solid #CCCCCC;'>"&dic(md)&"</td>"
result = result & "</tr>"
	total = total + rs("P_Price") * dic(md)
		rs.close:set rs = nothing
		session("total")=total
	Next
result = result & "<tr bgcolor='#FFFFFF' align='center'><td colspan='4' bgcolor='#FFFFFF' align='center'>总价格："&total&"元</td></tr>"

result = result & "</table>"
SelectProductList = result
End Function
'=======================================
'添加订单
'=======================================
Function GetRS(sql)

On Error Resume Next
Set GetRS = conn.exec(sql,"r1")
If Err Then
	'echo sql
	'die err.description
	die "错误，表结构不是最新版本"
End If
End Function



'=======================================
'订单打印
'=======================================
Sub OrderPrint
if needCheck then
if getForm("code","post")<>Session("Code") then alertMsgAndGo "验证码不正确","-1"
end if
dim sql		
dim id, sortID,SortAndID
	dim qs
	qs = Session("Cart_QS")
	SortAndID=split(replaceStr(qs,FileExt,""),"_")
	
	if isNul(replaceStr(qs,FileExt,"")) then  echoMsgAndGo "页面不存在",3 
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "页面不存在",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "页面不存在",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "请选择产品！","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "请选择产品！","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productprint.html"
	if not CheckTemplateFile(templatePath) then echo "productprint.html"&err_16
	
	set rsObj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then echoMsgAndGo "页面不存在",3 : exit sub		
	with templateObj 
	
		.content=loadFile(templatePath)	
		.parseHtml()
		


'订单号码：{aspcms:order.no} 
templateObj.content=replace(templateObj.content,"{aspcms:order.no}",Session("Print.OrderNo"))
'订单日期：{aspcms:order.date} 
templateObj.content=replace(templateObj.content,"{aspcms:order.date}",Session("Print.OrderDate"))
'用 户 名：{aspcms:order.username} 
templateObj.content=replace(templateObj.content,"{aspcms:order.username}",Session("Print.OrderUserName"))
'订单状态：{aspcms:order.state} 
templateObj.content=replace(templateObj.content,"{aspcms:order.state}",Session("Print.OrderState"))
'订单总金额：{aspcms:order.total}元 
templateObj.content=replace(templateObj.content,"{aspcms:order.total}",Session("Print.OrderTotal"))
'商品总数量：{aspcms:order.count} 	
templateObj.content=replace(templateObj.content,"{aspcms:order.count}",Session("Print.OrderCount"))
'客户姓名：{aspcms:order.nicename}
templateObj.content=replace(templateObj.content,"{aspcms:order.nicename}",Session("Print.OrderNiceName"))
'联系电话：{aspcms:order.tel}
templateObj.content=replace(templateObj.content,"{aspcms:order.tel}",Session("Print.OrderTel"))
'手机号码：{aspcms:order.cellphone}
templateObj.content=replace(templateObj.content,"{aspcms:order.cellphone}",Session("Print.OrderCellphone"))
'邮政编码：{aspcms:order.zipcode}
templateObj.content=replace(templateObj.content,"{aspcms:order.zipcode}",Session("Print.OrderZipcode"))
'联系地址：{aspcms:order.address}
templateObj.content=replace(templateObj.content,"{aspcms:order.address}",Session("Print.OrderAddress"))
'电子邮箱：{aspcms:order.email}
templateObj.content=replace(templateObj.content,"{aspcms:order.email}",Session("Print.OrderEmail"))
'简单留言：{aspcms:order.note}
templateObj.content=replace(templateObj.content,"{aspcms:order.note}",Session("Print.OrderNote"))
'支付方式
templateObj.content=replace(templateObj.content,"{aspcms:order.payment}",Session("Print.OrderPayMent"))

select case Session("Print.OrderPayMent")
	case "在线支付"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Online)
	case "银行支付"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Bank)
	case "邮局付款"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_PostOffice)
end select


'商品列表
templateObj.content=replace(templateObj.content,"{aspcms:selectproduct}",SelectProductList)
 

		.parseCommon() 		
		echo .content 
	end with
	
	set templateobj =nothing : terminateAllObjects
End Sub



'=======================================
'订单确认
'=======================================
Sub OrderComfirm()
if needCheck then
if getForm("code","post")<>Session("Code") then alertMsgAndGo "验证码不正确","-1"
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
	
if isnul(nicename) then alertMsgAndGo "联系人不能为空","-1"
if isnul(tel) then alertMsgAndGo "联系电话不能为空","-1"
if isnul(cellphone) then alertMsgAndGo "手机不能为空","-1"
		
		
dim id, sortID,SortAndID
	dim qs
	qs = Session("Cart_QS")
	SortAndID=split(replaceStr(qs,FileExt,""),"_")
	
	if isNul(replaceStr(qs,FileExt,"")) then  echoMsgAndGo "页面不存在",3 
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "页面不存在",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "页面不存在",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "请选择产品！","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "请选择产品！","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productcomfirm.html"
	if not CheckTemplateFile(templatePath) then echo "productcomfirm.html"&err_16
	
	set rsObj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then echoMsgAndGo "页面不存在",3 : exit sub		
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
hidden = hidden & "<input type='submit' value='确认订单' name='btnSubmit'/></p></form>"

'订单参数传递 {aspcms:order.$trigger$}
templateObj.content=replace(templateObj.content,"{aspcms:order.$trigger$}",hidden)

If m_username = "" then m_username = "游客"

'订单号码：{aspcms:order.no} 
templateObj.content=replace(templateObj.content,"{aspcms:order.no}",orderno)
Session("Print.OrderNo") = orderno
'订单日期：{aspcms:order.date} 
templateObj.content=replace(templateObj.content,"{aspcms:order.date}",addtime)
Session("Print.OrderDate") = addtime
'用 户 名：{aspcms:order.username} 
templateObj.content=replace(templateObj.content,"{aspcms:order.username}",m_username)
Session("Print.OrderUserName") = m_username
'订单状态：{aspcms:order.state} 
templateObj.content=replace(templateObj.content,"{aspcms:order.state}","未处理")
Session("Print.OrderState") = "未处理"
'订单总金额：{aspcms:order.total}元 
templateObj.content=replace(templateObj.content,"{aspcms:order.total}",oTotal)
Session("Print.OrderTotal") = oTotal
'商品总数量：{aspcms:order.count} 	
templateObj.content=replace(templateObj.content,"{aspcms:order.count}",oCount)
Session("Print.OrderCount") = oCount
'客户姓名：{aspcms:order.nicename}
templateObj.content=replace(templateObj.content,"{aspcms:order.nicename}",nicename)
Session("Print.OrderNiceName") = nicename
'联系电话：{aspcms:order.tel}
templateObj.content=replace(templateObj.content,"{aspcms:order.tel}",tel)
Session("Print.OrderTel") = tel
'手机号码：{aspcms:order.cellphone}
templateObj.content=replace(templateObj.content,"{aspcms:order.cellphone}",cellphone)
Session("Print.OrderCellphone") = cellphone
'邮政编码：{aspcms:order.zipcode}
templateObj.content=replace(templateObj.content,"{aspcms:order.zipcode}",zipcode)
Session("Print.OrderZipcode") = zipcode
'联系地址：{aspcms:order.address}
templateObj.content=replace(templateObj.content,"{aspcms:order.address}",address)
Session("Print.OrderAddress") = address
'电子邮箱：{aspcms:order.email}
templateObj.content=replace(templateObj.content,"{aspcms:order.email}",email)
Session("Print.OrderEmail") = email
'简单留言：{aspcms:order.note}
templateObj.content=replace(templateObj.content,"{aspcms:order.note}",note)
Session("Print.OrderNote") = note
'支付方式
templateObj.content=replace(templateObj.content,"{aspcms:order.payment}",Payment)
Session("Print.OrderPayMent") = Payment

select case Payment
	case "在线支付"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Online)
	case "银行支付"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_Bank)
	case "邮局付款"
	templateObj.content=replace(templateObj.content,"{aspcms:order.paymentdesc}",Payment_PostOffice)
end select


'商品列表
templateObj.content=replace(templateObj.content,"{aspcms:selectproduct}",SelectProductList)
 

		.parseCommon() 		
		echo .content 
	end with
	
	set templateobj =nothing : terminateAllObjects
	
End Sub
%>
