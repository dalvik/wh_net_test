<!--#include file="AspCms_OrderFun.asp" -->
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
Dim m_orderno,sql,rs

m_orderno = getForm("OrderNo","both")

sql = "select {prefix}Order2.*,(select sum(count) from {prefix}OrderProduct where {prefix}OrderProduct.orderno={prefix}Order2.orderno) as pcount,(select sum(InstantPrice*[Count]) from {prefix}OrderProduct where {prefix}OrderProduct.orderno={prefix}Order2.orderno) as pprice from {prefix}Order2 where orderno = '"&m_orderno&"'"
Set rs = Conn.exec(sql,"r1")

if rs.eof or rs.bof then 
	die "δ�ҵ�������" & m_orderno
end if

%>
</head>
<body>
<div class="oc">
<div class="oc_thead"><div>������Ϣ</div></div>
<div class="oc_content"> 
<table cellpadding="0" cellspacing="25" border="0" align="center" width="90%">
  <tr align="center">
    <td width="50%" align="left">
	<p>�������룺<%=m_orderno%>
  <p>�������ڣ�<%=rs("OrderTime")%>
  <p>�� �� ����<%=rs("UserName")%>
  <p>����״̬��<%if rs("State") = 0 then response.write "δ����" else response.write "�Ѵ���" end if%>
  <p>��Ʒ��������<%=rs("pcount")%>
  <p>�����ܽ�<%=rs("pprice")%>Ԫ  
  <p>������ע��<br>
  * ���ס��������.<br>
  * ֧��ȷ�Ϻ�,�����յ�һ��֧��(����)�ĵ����ʼ� 
	</td>
    <td>
    
<table width="245" cellpadding="0" cellspacing="0">
<tr>
<td><fieldset style="width:245px" align="center"> <legend> ��������Ϣ </legend>
<p>�ͻ�������<%=rs("nicename")%></p>
<p>��ϵ�绰��<%=rs("tel")%></p>
<p>�ֻ����룺<%=rs("cellphone")%></p>
<p>�������룺<%=rs("zipcode")%></p>
<p>��ϵ��ַ��<%=rs("address")%></p>
<p>�������䣺<%=rs("email")%></p>
<p>�����ԣ�<%=rs("note")%></p>
</fieldset></td>
</tr>
</table>
</td>
  </tr>
</table>
</div>


<div class="oc_thead"><div>������Ϣ</div></div>
<div class="oc_content">
<table width='80%' border='0' cellspacing='1' cellpadding='0' bgcolor='#cccccc' align="center">
<tr bgcolor='#FFFFFF' align='center'>
<td bgcolor='#FFFFFF'>��ƷͼƬ</td>
<td bgcolor='#FFFFFF'>��Ʒ����</td>
<td bgcolor='#FFFFFF'>����</td>
<td bgcolor='#FFFFFF'>����</td>
</tr>
<%
Dim prs
sql = "select op.count,op.InstantPrice,op.productid,(select title from {prefix}content where contentid=op.productid) as title,(select IndexImage from {prefix}content where contentid=op.productid) as IndexImage from {prefix}OrderProduct op where orderno='"&m_orderno&"'"
set prs = conn.exec(sql,"r1")
while not prs.eof
%>
<tr bgcolor='#FFFFFF' align='center'>
<td bgcolor='#FFFFFF'><a title='����鿴ԭͼ' href='../../../content/?<%=prs("productid")%>.html'  target='_blank'><img src='<%=prs("IndexImage")%>' width='70' height='50' /></a></td>
<td bgcolor='#FFFFFF'><a title='<%=prs("title")%>' href='../../../content/?<%=prs("productid")%>.html'  target=_blank><%=prs("title")%></a></td>
<td bgcolor='#FFFFFF'><%=prs("InstantPrice")%>Ԫ</td>
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


<div class="oc_thead"><div>�˵���ϸ</div></div>
<div class="oc_content">
<dl>
<dd>�� �� � <%=rs("pprice")%>Ԫ(��֧�����)
</dl>
</ul>
</div>



<div class="oc_thead"><div>֧����ʽ</div></div>
<div class="oc_content oc_content_b">
<p style="padding-left:40px"><%=rs("payment")%></p>
<p style="padding-left:40px">
<%
select case rs("payment")
	case "����֧��"
	echo Payment_Online
	case "����֧��"
	echo Payment_Bank
	case "�ʾָ���"
	echo Payment_PostOffice
end select
%>


</p>
</div>
</div>
</body>
</html><%rs.close:set rs = nothing%>