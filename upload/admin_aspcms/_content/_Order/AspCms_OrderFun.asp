<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<%
CheckAdmin("AspCms_Order.asp")
'定义类别ID，搜索关键词，页数，排序


dim sortType, keyword, page, psize, order, ordsc, sortid, MOrderNo
sortType=getForm("sortType","get")
if isnul(sortType) then sortType=0	
sortid=getForm("sortid","post")	
if isnul(sortid) then sortid=getForm("sortid","get")
keyword=getForm("keyword","post")
if isnul(keyword) then keyword=getForm("keyword","get")
page=getForm("page","get")
psize=getForm("psize","get")
order=getForm("order","get")
ordsc=getForm("ordsc","get")
 


dim action : action=getForm("action","get")

select case action 
	case "edit" : editOrder	
	case "del" : delOrder	
	case "on" : SetState'onOff "on", "Order", "OrderID", "OrderStatus", "", getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc
	case "off" : SetState'onOff "off", "Order", "OrderID", "OrderStatus", "", getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc
end select


Dim OrderID,UserID,OrderName,Linkman,Gender,Address,PostCode,Phone,Mobile,Email,QQ,AddTime,Remark,OrderStatus,Num

Sub SetState()
dim orderno,s
dim sql

orderno = getForm("orderno","get")
s = getForm("s","get")

sql = "update {prefix}Order2 set state = " & s & " where orderno='"&orderno&"'"
	call conn.exec(sql,"exe")
End Sub

Sub delOrder
	dim sql
	Dim id	:	id=getForm("id","both")
	'die "删除" &id
	if isnul(id) then alertMsgAndGo "请选择要删除的内容","-1"
	id = "'" & id & "'"
	id = replace(id,",","','")
	id = replace(id," ","")
	sql = "delete from {prefix}Order2 where OrderNo in("&id&")"
	Conn.Exec sql,"exe"
	'echo sql & "<br>"
	sql = "delete from {prefix}OrderProduct where OrderNo in("&id&")"
	Conn.Exec sql,"exe"
	'echo sql & "<br>"
	alertMsgAndGo "删除成功","?page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword
End Sub

Sub getContent

	dim id : id=getForm("id","get")
	if not isnul(ID) then		
		conn.exec"update {prefix}Order set OrderStatus=1 Where OrderID="&id,"exe"
		Dim rs : Set rs = Conn.Exec("select * from {prefix}Order where OrderID="&ID,"r1")
		if not rs.eof then	
			OrderID=rs("OrderID")
			UserID=rs("UserID")
			OrderName=rs("OrderName")
			Linkman=rs("Linkman")
			Gender=rs("Gender")
			Address=rs("Address")
			PostCode=rs("PostCode")
			Phone=rs("Phone")
			Mobile=rs("Mobile")
			Email=rs("Email")
			QQ=rs("QQ")
			Num=rs("Num")
			AddTime=rs("AddTime")
			Remark=rs("Remark")
			OrderStatus=rs("OrderStatus")							
		end if
	else		
		alertMsgAndGo "没有这条记录","-1"
	end if
End Sub

Sub editOrder
	OrderID=getForm("OrderID","post")	
	UserID=filterPara(getForm("UserID","post"))
	Num=filterPara(getForm("Num","post"))
	OrderName=filterPara(getForm("OrderName","post"))
	Linkman=filterPara(getForm("Linkman","post"))
	Gender=filterPara(getForm("Gender","post"))
	Address=filterPara(getForm("Address","post"))
	PostCode=filterPara(getForm("PostCode","post"))
	Phone=filterPara(getForm("Phone","post"))
	Mobile=filterPara(getForm("Mobile","post"))	
	Email=filterPara(getForm("Email","post"))
	QQ=filterPara(getForm("QQ","post"))	
	'AddTime=now()
	OrderStatus=getCheck(getForm("OrderStatus","post"))	
	
	conn.exec "update {prefix}Order set OrderName='"&OrderName&"',Num="&Num&",Linkman='"&Linkman&"',Gender="&Gender&",Address='"&Address&"',PostCode='"&PostCode&"',Phone='"&Phone&"',Mobile='"&Mobile&"',Email='"&Email&"',QQ='"&QQ&"',OrderStatus="&OrderStatus&"  where OrderID="&OrderID,"exe"
	alertMsgAndGo "修改成功","AspCms_Order.asp?page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword
End Sub

Sub OrderList	
	dim datalistObj,rsArray
	dim m,i,orderStr,whereStr,sqlStr,rsObj,allPage,allRecordset,numPerPage,searchStr
	numPerPage=10
	orderStr= " order by ordertime desc"
	if isNul(page) then page=1 else page=clng(page)
	if page=0 then page=1
	whereStr=" where 1=1 "
	if not isNul(SortID) then  whereStr=whereStr
	if not isNul(keyword) then 
		whereStr = whereStr&" and nicename like '%"&keyword&"%' or tel like '%"&keyword&"%' or cellphone like '%"&keyword&"%'"
	end if
	
	

	sqlStr = "select OrderNo,UserName,OrderTime,State,Payment,nicename,tel,cellphone,zipcode,address,email,note,to,Invoice,(select sum(instantprice*count) from {prefix}OrderProduct where {prefix}OrderProduct.OrderNo = {prefix}Order2.OrderNo) as totalprice from {prefix}Order2 "&whereStr&orderStr
	'die sqlstr
	set rsObj = conn.Exec(sqlStr,"r1")
	rsObj.pagesize = numPerPage
	allRecordset = rsObj.recordcount : allPage= rsObj.pagecount
	if page>allPage then page=allPage
	if allRecordset=0 then
		if not isNul(keyword) then
		    echo "<tr bgcolor=""#FFFFFF"" align=""center""><td colspan='9'>关键字 <font color=red>"""&keyword&"""</font> 没有记录</td></tr>" 
		else
		    echo "<tr bgcolor=""#FFFFFF"" align=""center""><td colspan='9'>还没有记录!</td></tr>"
		end if
	else  
'选择
'订单编号
'订购人
'联系电话
'总金额
'订购时间
'状态
'操作
		rsObj.absolutepage = page
		for i = 1 to numPerPage	
			 echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"
				echo"<td><input type=""checkbox"" name=""id"" value="""&rsObj("OrderNo")&""" class=""checkbox"" /></td>"
				echo"<td><a href=""AspCms_OrderDetail.asp?id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" >"&rsObj("OrderNo")&"</a></td>"
				'echo"<td>"&rsObj("UserName")&"</td>"
				echo"<td>"&rsObj("NiceName")&"</td>"
				echo"<td>"&rsObj("cellphone")&"</td>"
				echo"<td>"&rsObj("totalprice")&"</td>"
				echo"<td>"&rsObj("OrderTime")&"</td>"
								
				'echo "<td>"&getStr(rsObj("State"),"<a href=""?action=off&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" title=""已处理"" ><IMG src=""../../images/toolbar_ok.gif""></a>","<a href=""?action=on&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" title=""未处理"" ><IMG src=""../../images/toolbar_no.gif""></a>")&"</td>"
				
				echo "<td>"&getStr(rsObj("State"),"<a href=""?action=off&orderno="&rsObj(0)&"&s=0"" title=""已处理"" ><IMG src=""../../images/toolbar_ok.gif""></a>","<a href=""?action=on&orderno="&rsObj(0)&"&s=1"" title=""未处理"" ><IMG src=""../../images/toolbar_no.gif""></a>")&"</td>"
				
				echo "<td><a href='AspCms_OrderInfo.asp?orderno="&rsObj(0)&"' target='_blank'>订单信息</a>"
				echo " | "
				'echo "<a href=""AspCms_OrderDetail.asp?id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" >查看</a>"
				'echo " | "
				echo "<a href=""?action=del&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" onClick=""return confirm('确定要删除吗')"">删除</a>"
				echo "</td>"
				
			  echo "</tr>"&vbcrlf
			rsObj.movenext
			if rsObj.eof then exit for
		next

		echo"<tr bgcolor=""#FFFFFF"" class=""pagenavi"">"
		echo"<td colspan=""8"" height=""28"" style=""padding-left:20px;"">"
		echo"页数："&page&"/"&allPage&"  每页"&numPerPage &" 总记录数"&allRecordset&"条 <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page=1&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">首页</a> <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&(page-1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">上一页</a> "&vbcrlf
		dim pageNumber
		pageNumber=makePageNumber_(page, 10, allPage, "newslist",sortID, order,keyword)
		echo pageNumber
		echo"<a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&(page+1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">下一页</a> <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&allPage&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">尾页</a>"&vbcrlf&_		
			"</td>"&vbcrlf& _			
		"</tr>"&vbcrlf
	end if
	rsObj.close : set rsObj = nothing	
End Sub





Sub OrderListD
dim id :id =getForm("id","get")
	dim datalistObj,rsArray
	dim m,i,orderStr,whereStr,sqlStr,rsObj,allPage,allRecordset,numPerPage,searchStr
	numPerPage=10
	orderStr= " order by id desc"
	if isNul(page) then page=1 else page=clng(page)
	if page=0 then page=1
	whereStr=" where orderno='"&id&"' "
	if not isNul(SortID) then  whereStr=whereStr
	if not isNul(keyword) then 
		'whereStr = whereStr&" and nicename like '%"&keyword&"%' or tel like '%"&keyword&"%' or cellphone like '%"&keyword&"%'"
	end if
	sqlStr = "select orderno,(select title from {prefix}Content where contentid=op.ProductID) as productname,(select p_price from {prefix}Content where contentid=op.ProductID) as productprice,count,instantprice from {prefix}OrderProduct as op "&whereStr&orderStr
	
	set rsObj = conn.Exec(sqlStr,"r1")
	rsObj.pagesize = numPerPage
	allRecordset = rsObj.recordcount : allPage= rsObj.pagecount
	if page>allPage then page=allPage
	if allRecordset=0 then
		if not isNul(keyword) then
		    echo "<tr bgcolor=""#FFFFFF"" align=""center""><td colspan='9'>关键字 <font color=red>"""&keyword&"""</font> 没有记录</td></tr>" 
		else
		    echo "<tr bgcolor=""#FFFFFF"" align=""center""><td colspan='9'>还没有记录!</td></tr>"
		end if
	else  
'选择
'订单编号
'产品名称
'订购数量
'即时单价
'操作
		rsObj.absolutepage = page
		for i = 1 to numPerPage	
			 echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"
				echo "<td><input type=""checkbox"" name=""id"" value="""&rsObj("OrderNo")&""" class=""checkbox"" /></td>"
				echo "<td>"&rsObj("OrderNo")&"</td>"
				echo "<td>"&rsObj("productname")&"</td>"
				echo "<td>"&rsObj("productprice")&"</td>"
				echo "<td>"&rsObj("count")&"</td>"
				echo "<td>"&rsObj("instantprice")&"</td>"				
				'echo "<td>"&getStr(rsObj("State"),"<a href=""?action=off&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" title=""已查看"" ><IMG src=""../../images/toolbar_ok.gif""></a>","<a href=""?action=on&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" title=""未查看"" ><IMG src=""../../images/toolbar_no.gif""></a>")&"</td>"
				'echo "<td><a href=""AspCms_OrderEdit.asp?id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" >查看</a> | <a href=""?action=del&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&"""  onClick=""return confirm('确定要删除吗')"">删除</a></td>"&
				
			  echo "</tr>"&vbcrlf
			rsObj.movenext
			if rsObj.eof then exit for
		next

		echo"<tr bgcolor=""#FFFFFF"" class=""pagenavi"">"&vbcrlf& _
			"<td colspan=""8"" height=""28"" style=""padding-left:20px;"">"&vbcrlf& _			
		"页数："&page&"/"&allPage&"  每页"&numPerPage &" 总记录数"&allRecordset&"条 <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page=1&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">首页</a> <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&(page-1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">上一页</a> "&vbcrlf
		dim pageNumber
		pageNumber=makePageNumber_(page, 10, allPage, "newslist",sortID, order,keyword)
		echo pageNumber
		echo"<a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&(page+1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">下一页</a> <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&allPage&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">尾页</a>"&vbcrlf&_		
			"</td>"&vbcrlf& _			
		"</tr>"&vbcrlf
	end if
	rsObj.close : set rsObj = nothing	
End Sub


%>