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
    <TD height=30></TD>
    <TD align=right colSpan=2>
    <INPUT onClick="history.go(-1)" type="button" value="����" class="button" /> 
</TD></TR></TBODY></TABLE></DIV>
<DIV class=listzone>
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
  <TBODY>
  <TR class=list>
  
    <TD width=44 align="center" class=biaoti>ѡ��</TD>
    <TD width=110 align="center" class=biaoti>�������</TD>
    <TD width=* align="center" class=biaoti>��Ʒ����</TD>
    <TD width=110 align="center" class=biaoti>��Ʒ�۸�</TD>
    <TD width=189 align="center" class=biaoti>��������</TD>
    <TD width=110 align="center" class=biaoti><span class="searchzone">���ۼ۸�</span></TD>       
    <!--<TD width=90 align="center" class=biaoti><span class="searchzone">����</span></TD>-->
    </TR>
	<%OrderListD%>
    </TBODY></TABLE>
</DIV>
<DIV class="piliang">
<INPUT onClick="SelAll(this.form)" type="checkbox" value="1" name="SELALL"> ȫѡ&nbsp;
<!--<INPUT class="button" type="submit" value="ɾ��" onClick="if(confirm('ȷ��Ҫ����ɾ����?\n\n�˲���������!')){form.action='?action=del<%="&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>';}else{return false};"/>  -->



</DIV>
</FORM>
</BODY></HTML>
