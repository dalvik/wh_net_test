<!--#include file="AspCms_smssortFun.asp" -->
<%CheckLogin()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="images/style.css" type=text/css rel=stylesheet>
</HEAD>
<SCRIPT>
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
<!--#include file="AspCms_smsHead.asp" -->

<DIV class=searchzone>
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD height=30>&nbsp;</TD>
    <TD align=right colSpan=2>   <INPUT onClick="location.href='AspCms_smsnumAdd.asp'" type="button" value="��Ӻ���" class="button" > <INPUT onClick="location.href='AspCms_smssort.asp'" type="button" value="��Ӻ������" class="button" > <INPUT onClick="location.href='<%=getPageName()%>?<%="smsSortID="&smsSortID&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>'" type="button" value="ˢ��" class="button" /> </TD>
  </TR></TBODY></TABLE></DIV>
<FORM name="" action="?" method="post">
<DIV class=searchzone>

<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD width="85%" height=30><%loadsmsbutton%></TD>
    <TD width="15%" colSpan=2 align=right>
    
    
  
    
</TD></TR></TBODY></TABLE></DIV>
<DIV class=listzone>
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0>
  <TBODY>
  <TR class=list>
    <TD width=96 align="center" class=biaoti>ѡ��</TD>
    <TD width=94 align="center" class=biaoti>���</TD>
    <TD width="96" height=28 align="left" class=biaoti>����</TD>
    <TD width=178 align="center" class=biaoti><span class="searchzone">�ֻ���</span></TD>
    <TD width=150 align="center" class=biaoti><span class="searchzone">����</span></TD>
    <TD width=860 align="center" class=biaoti><span class="searchzone">��ע</span></TD>
    <TD width=267 align="center" class=biaoti><span class="searchzone">����</span></TD>
    </TR>
	<%smsnumlist%>
    </TBODY></TABLE>
</DIV>
<DIV class="piliang">
<INPUT onClick="SelAll(this.form)" type="checkbox" value="1" name="SELALL"> ȫѡ&nbsp;
<INPUT class="button" type="submit" value="ɾ��" onClick="if(confirm('ȷ��Ҫɾ����')){form.action='?action=del&<%="smsSortID="&smsSortID&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>';}else{return false};"/>  


Ŀ�����:<%loadsmsSelect 0%>
<INPUT class="button" type="submit" value="�ƶ�" onClick="form.action='?action=move&<%="smsSortID="&smsSortID&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>';"/> <INPUT class="button" type="submit" value="����" onClick="form.action='?action=copy&<%="smsSortID="&smsSortID&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""%>';"/>

</DIV>
</FORM>
</BODY></HTML>
