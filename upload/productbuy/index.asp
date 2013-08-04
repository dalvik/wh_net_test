<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
'echo "进入cart"
'echo getForm("proid","both")
'Session.Abandon()
dim action:action=getForm("act","both")
'dim proid:proid=getForm("id","both")
dim proid:proid=cint(request("proid"))
 

'Set Session("Cart") = Nothing
'Set Session("Cart.Price") = Nothing

Add2Session
Cart_Load


'加入购物车
Sub Add2Session
Dim dicCount,dicPrice,pid,count,price
pid = getForm("proid","post")
count = getForm("count","post")
price = getForm("proprice","post")
pid = cint(pid)
count = cint(count)

if pid="" then exit sub
If IsEmpty(Session("Cart")) Then
	Set dicCount = Server.CreateObject("Scripting.Dictionary")
	Set dicPrice = Server.CreateObject("Scripting.Dictionary")
Else
	Set dicCount = Session("Cart")
	Set dicPrice = Session("Cart.Price")
End If

if dicCount.Exists(pid) then
	dicCount(pid) = count
	dicPrice(pid) = price
else
	dicCount.add pid,count
	dicPrice.add pid,price
end if

Set Session("Cart") = dicCount
Set Session("Cart.Price") = dicPrice
End Sub
'所有保存
Sub All2Cart
Dim dicCount,dicPrice,pid,count,price,i,keysPrice,keysCount
'pid = getForm("proid","post")
'count = getForm("count","post")
'price = getForm("price","post")
if pid="" then exit sub
If IsEmpty(Session("Cart")) Then
	Set dicCount = Server.CreateObject("Scripting.Dictionary")
	Set dicPrice = Server.CreateObject("Scripting.Dictionary")
Else
	Set dicCount = Session("Cart")
	Set dicPrice = Session("Cart.Price")
End If
keysPrice = dicPrice.keys
keysCount = dicCount.keys
for i=0 to dicCount.Count-1
	if dicCount.Exists(keysCount(i)) then
		echo getForm("count"&keysCount(i),"post")
		dicCount(keysCount(i)) = getForm("count"&keysCount(i),"post")
		echo getForm("proprice"&keysPrice(i),"post")
		dicPrice(keysPrice(i)) = getForm("proprice"&keysPrice(i),"post")
	else
		dicCount.add keysCount(i),getForm("count"&keysCount(i),"post")
		dicPrice.add keysPrice(i),getForm("proprice"&keysPrice(i),"post")
	end if
next
Set Session("Cart") = dicCount
Set Session("Cart.Price") = dicPrice
End Sub




'从购物车删除
Sub DeleteFromSession(proid)
Dim dicCount,dicPrice
	If Not IsEmpty(Session("Cart")) Then
		Set dicCount = Session("Cart")
		Set dicPrice = Session("Cart.Price")
		If dicCount.exists(proid) Then dic.remove(proid)
		If dicPrice.exists(proid) Then dic.remove(proid)
	End If	
	Set Session("Cart") = dicCount
	Set Session("Cart.Price") = dicPrice
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub



'显示购物车
Function CartList
Dim dic,md
Dim rs,sql
Dim result,total,tcount
Dim fid

fid=getForm("proid","post")


result = "<form name='frmCart' action='?action=all2cart'>"
result = result & "<ul class=""atable"">"
result = result & "<li class=""ahead"">"
result = result & "<div class=""p_img"">商品图片</div>"
result = result & "<div class=""p_name"">商品名称</div>"
result = result & "<div class=""p_price"">价格</div>"
result = result & "<div class=""p_count"">数量</div>"
result = result & "<div class=""p_total"">小计</div>"
result = result & "<div>操作</div>"
'<input type='hidden' name='action' value='all2cart' />
result = result & "</li>"
tcount = 0

If Not IsEmpty(Session("Cart")) Then
	Set dic = Session("Cart")
	For Each md In dic
		'echo md & dic(md) & "<hr>"
		sql = "select P_Price,IndexImage,title,contentid from {prefix}Content where contentid=" & md
		Set rs = GetRS(sql)
result = result & "<li>"
result = result & "<div class='p_img'><a title='点击查看原图' href='../content/?"&rs("contentid")&".html'  target=_blank>"
if rs("IndexImage") = "" then
result = result & "<img src='../Images/nopic.gif' /></a></div>"
else
result = result & "<img src='"&rs("IndexImage")&"' /></a></div>"
end if

result = result & "<div class='p_name'><a title='"&rs("title")&"' href='../content/?"&rs("contentid")&".html'  target=_blank>"&rs("title")&"</a></div>"
result = result & "<div class='p_price'><span id='price"&md&"'>"&rs("P_Price")&"</span>元<input type='hidden' name='price"&md&"' value='"&rs("P_Price")&"'></div>"
result = result & "<div class='p_count'><input id='count"&md&"' name='count"&md&"' maxlength='5' onkeyup=""this.value=this.value.replace(/\D/g,'');changePrice(this,"&md&")""  onafterpaste=""this.value=this.value.replace(/\D/g,'')"" onblur=""if(value==''){value='1';changePrice(this,"&md&")}"" maxlength='5' value='"&dic(md)&"'"
result = result & "'/></div>"
result = result & "<div class='p_total'><span id='curtotal"&md&"'>"&dic(md) * rs("P_Price")&"</span>元</div>"
result = result & "<div><a href='javascript:void(0)' onclick='deleteItem("&md&",this)'>删除</a></div>"
result = result & "</li>"
total = total + (dic(md) * rs("P_Price"))
tcount = tcount + dic(md)
rs.close:set rs = nothing
	Next
	
End If
result = result & "</ul>"
result = result & "<div style='float:left;width:600px'>你的购物车里有商品：<span id='allcount'>"&tcount&"</span> 件 共计：<span id='alltotal'>"&total&"</span> </div>"
result = result & "<div style='float:left;width:600px'><a class='button_2' style='float:left;letter-spacing:0px; margin-right:20px;' href='../list/?5_1.html'>继续购物</a> <a class='button_2' style='float:left' href='checkout.asp'>订购</a></div>"
result = result & "</form>"

	CartList = result
	'die result
End Function


Function GetRS(sql)

On Error Resume Next
Set GetRS = conn.exec(sql,"r1")
If Err Then
	'echo sql
	'die err.description
	die "错误，表结构不是最新版本"
End If
End Function


Sub TestA
	Dim dic,md
	Set dic = Session("Cart")

	for each md in dic
		echo md & dic(md) & "<hr>"  
	next

	die "结束"	
End Sub

Sub Cart_Load	
	dim id, sortID,SortAndID
	SortAndID=split(replaceStr(request.QueryString,FileExt,""),"_")
	if isNul(replaceStr(request.QueryString,FileExt,"")) then  echoMsgAndGo "页面不存在",3 
	Session("Cart_QS") = Request.QueryString
	SortID = SortAndID(0)
	if not isNul(SortID) and isNum(SortID) then SortID=clng(SortID) else echoMsgAndGo "页面不存在",3 end if
	id=SortAndID(1)
	if not isNul(id) and isNum(id) then id=clng(id) else echoMsgAndGo "页面不存在",3 end if
	
	if isnul(id) or not isnum(id) then alertMsgAndGo "请选择产品！","-1" 
	if isnul(sortID) or not isnum(sortID) then alertMsgAndGo "请选择产品！","-1" 
	
	dim templateobj,channelTemplatePath : set templateobj = new TemplateClass	
	dim typeIds,rsObj,rsObjtid,Tid,rsObjSmalltype,rsObjBigtype
	Dim templatePath,tempStr
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/productcart.html"
	if not CheckTemplateFile(templatePath) then echo "productcart.html"&err_16
	
	set rsobj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then 
	echoMsgAndGo "页面不存在",3
	'exit sub		
	end if
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
		
	dim selectproduct	
	set rsObj=conn.Exec("select title from {prefix}Content where ContentID="&id,"r1")
	selectproduct=rsObj(0)
	
	rsObj.close()		
		.content=replaceStr(.content,"{aspcms:cartlist}",CartList)
		.parseCommon() 
		echo .content 
	end with
	set templateobj =nothing : terminateAllObjects
End Sub

%>