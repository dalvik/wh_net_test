<!--#include file="../inc/AspCms_SettingClass.asp" -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en" "http://www.w3c.org/tr/1999/rec-html401-19991224/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head><title></title>
<meta http-equiv=content-type content="text/html; charset=gbk">

<style type="text/css">
.oc{min-width:583px;border:#cccccc 1px solid;}
.oc_thead{background:url(/templates/cn/images/ocTHead.gif); height:26px; line-height:26px}
.oc_thead div{background:url(/templates/cn/images/bg.png) no-repeat left -43px;height:26px; width:120px; text-align:center; color:white; font-weight:bold;float:left}
.oc_content{ border:0px; margin:0px; padding:5px; }
.oc_content_b{border-bottom:#cccccc 1px solid;}
.oc_content_h{border-top:#cccccc 1px solid;}
fieldset p{ text-align:left; margin:0px; padding:3px;}
img{border:0px}


.bank{ list-style:none}
.banklogo{ border:0px;}
</style>
<%
dim userID
userid=session("userID")
if isNul(userID) or not isNum(userID) then alertMsgAndGo "对不起,您的登录状态已经失效,请重新登录!","login.asp"
Dim m_orderno,sql,rs

m_orderno = getForm("OrderNo","both")

sql = "select {prefix}Order2.*,(select sum(count) from {prefix}OrderProduct where {prefix}OrderProduct.orderno={prefix}Order2.orderno) as pcount,(select sum(InstantPrice*[Count]) from {prefix}OrderProduct where {prefix}OrderProduct.orderno={prefix}Order2.orderno) as pprice from {prefix}Order2 where orderno = '"&m_orderno&"'"
Set rs = Conn.exec(sql,"r1")

if rs.eof or rs.bof then 
	die "未找到订单号" & m_orderno
end if

%>
</head>
<body>
<div class="oc">
<div class="oc_thead"><div>订单信息</div></div>
<div class="oc_content"> 
<table cellpadding="0" cellspacing="25" border="0" align="center" width="90%">
  <tr align="center">
    <td width="50%" align="left">
	<p>订单号码：<%=m_orderno%>
  <p>订单日期：<%=rs("OrderTime")%>
  <p>用 户 名：<%=rs("UserName")%>
  <p>订单状态：<%if rs("State") = 0 then response.write "未处理" else response.write "已处理" end if%>
  <p>商品总数量：<%=rs("pcount")%>
  <p>订单总金额：<%=rs("pprice")%>元  
  <p>订单备注：<br>
  * 请记住订单号码.<br>
  * 支付确认后,您将收到一封支付(结算)的电子邮件 
	</td>
    <td>
    
<table width="245" cellpadding="0" cellspacing="0">
<tr>
<td><fieldset style="width:245px" align="center"> <legend> 订购人信息 </legend>
<p>客户姓名：<%=rs("nicename")%></p>
<p>联系电话：<%=rs("tel")%></p>
<p>手机号码：<%=rs("cellphone")%></p>
<p>邮政编码：<%=rs("zipcode")%></p>
<p>联系地址：<%=rs("address")%></p>
<p>电子邮箱：<%=rs("email")%></p>
<p>简单留言：<%=rs("note")%></p>
</fieldset></td>
</tr>
</table>
</td>
  </tr>
</table>
</div>


<div class="oc_thead"><div>购物信息</div></div>
<div class="oc_content">
<table width='100%' border='0' cellspacing='1' cellpadding='0' bgcolor='#CCCCCC'>
<tr bgcolor='#FFFFFF' align='center'>
<td bgcolor='#FFFFFF'>商品图片</td>
<td bgcolor='#FFFFFF'>商品名称</td>
<td bgcolor='#FFFFFF'>单价</td>
<td bgcolor='#FFFFFF'>数量</td>
</tr>
<%
Dim prs
sql = "select op.count,op.InstantPrice,op.productid,(select title from {prefix}content where contentid=op.productid) as title,(select IndexImage from {prefix}content where contentid=op.productid) as IndexImage from {prefix}OrderProduct op where orderno='"&m_orderno&"'"
set prs = conn.exec(sql,"r1")
while not prs.eof
%>
<tr bgcolor='#FFFFFF' align='center'>
<td bgcolor='#FFFFFF'><a title='点击查看原图' href='../../../content/?<%=prs("productid")%>.html'  target='_blank'><img src='<%=prs("IndexImage")%>' width='70' height='50' /></a></td>
<td bgcolor='#FFFFFF'><a title='<%=prs("title")%>' href='../../../content/?<%=prs("productid")%>.html'  target=_blank><%=prs("title")%></a></td>
<td bgcolor='#FFFFFF'><%=prs("InstantPrice")%>元</td>
<td bgcolor='#FFFFFF'><%=prs("count")%></td>
</tr>
<%
prs.movenext
wend
prs.close
set prs = nothing
%>

</table>
</div>


<div class="oc_thead"><div>账单明细</div></div>
<div class="oc_content">
<dl>
<dd>总 金 额： <%=rs("pprice")%>元(需支付金额)
</dl>
</ul>
</div>



<div class="oc_thead"><div>支付方式</div></div>
<div class="oc_content oc_content_b">
<p style="padding-left:40px"><%=rs("payment")%></p>
<p style="padding-left:40px">
<%
select case rs("payment")
	case "在线支付"
	echo Payment_Online
	case "银行支付"
	echo Payment_Bank
	case "邮局付款"
	echo Payment_PostOffice
end select
%>


</p>
</div>
</div>
</body>
</html><%rs.close:set rs = nothing%>