<!--#include file="AspCms_OrderFun.asp" -->
<!doctype html public "-//w3c//dtd html 4.01 transitional//en" "http://www.w3c.org/tr/1999/rec-html401-19991224/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head><title></title>
<meta http-equiv=content-type content="text/html; charset=gbk">
<link href="../../images/style.css" type=text/css rel=stylesheet>
</head>
<script>
function SelAll(theForm){
		for ( i = 0 ; i < theForm.elements.length ; i ++ )
		{
			if ( theForm.elements[i].type == "checkbox" && theForm.elements[i].name != "SELALL" )
			{
				theForm.elements[i].checked = ! theForm.elements[i].checked ;
			}
		}
}
</SCRIPT>
<BODY>
<form action="?<%="sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>" method="post" name="form">
<DIV class=searchzone>

<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD height=30>��ϵ�˻���ϵ�绰��<input type="text" name="keyword" value="<%=keyword%>" class="input"/>&nbsp;&nbsp;  
      <INPUT class="button" type="submit" value="����" name="Submit" onClick="form.action='?<%="sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page=&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>';">
      <INPUT onClick="location.href='<%=getPageName()%>?<%="sortType="&sortType%>'" type="button" value="ȫ��" class="button" /></TD>
    <TD align=right colSpan=2>
    <INPUT onClick="location.href='<%=getPageName()%>?<%="sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>'" type="button" value="ˢ��" class="button" /> 
</TD></TR></TBODY></TABLE></DIV>
<DIV class=listzone>
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
  <TBODY>
  <TR class=list>
  
    <TD width=44 align="center" class=biaoti>ѡ��</TD>
    <TD width=110 align="center" class=biaoti>�������</TD>
    <TD width=110 align="center" class=biaoti><span class="searchzone">������</span></TD>
    <TD width=149 align="center" class=biaoti><span class="searchzone">��ϵ�绰</span></TD>
    <TD width=50 align="center" class=biaoti><span class="searchzone">�ܽ��</span></TD>
    <TD width=140 align="center" class=biaoti><span class="searchzone">����ʱ��</span></TD>
    <TD width=60 align="center" class=biaoti><span class="searchzone">֧��״̬</span></TD>
    <TD width=130 align="center" class=biaoti><span class="searchzone">����</span></TD>
    </TR>
	<%OrderList%>
    </TBODY></TABLE>
</DIV>
<DIV class="piliang">
<INPUT onClick="SelAll(this.form)" type="checkbox" value="1" name="SELALL"> ȫѡ&nbsp;
<INPUT class="button" type="submit" value="ɾ��" onClick="if(confirm('ȷ��Ҫ����ɾ����?\n\n�˲���������!')){form.action='?action=del<%="&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>';}else{return false};"/>  



</DIV>
</FORM>
</BODY></HTML>
