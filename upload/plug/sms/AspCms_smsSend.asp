<!--#include file="AspCms_smsSortFun.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="images/style.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.STYLE1 {font-weight: bold}
-->
</style>
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


<script language="javascript">

function getObject(objectId) {
 if(document.getElementById && document.getElementById(objectId)) {
 // W3C DOM
 return document.getElementById(objectId);
 }
 else if (document.all && document.all(objectId)) {
 // MSIE 4 DOM
 return document.all(objectId);
 }
 else if (document.layers && document.layers[objectId]) {
 // NN 4 DOM.. note: this won't find nested layers
 return document.layers[objectId];
 }
 else {
 return false;
 }
}


function blockHide(objname){
    var obj = getObject(objname);

		obj.style.display = "block";
		

}

function noneHide(objname){
    var obj = getObject(objname);
		obj.style.display = "none";
	
	

}
</script>
<BODY>
<!--#include file="AspCms_smsHead.asp" -->
<DIV class=searchzone>
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD height=30>
    	<span class="STYLE1">
        	<a href="#hsendname" target="_self">���û�����</a>&nbsp; 
            <a href="#hsendsort" target="_self">�����鷢��</a>&nbsp;
            <a href="#balance" target="_self">ʣ��ɷ�������</a>&nbsp;
            <a href="#hsendrecord" target="_self">���ͼ�¼</a>&nbsp;
        </span>
    </TD>
    <TD align=right colSpan=2>&nbsp;</TD>
  </TR></TBODY></TABLE></DIV>
<DIV class=tablezone id=hsendname><form name="sendname" method="post" action="">
  <TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
  <TBODY>
    <TR>
      <TD width="288" class=innerbiaoti><STRONG>���ѡȡȺ�����Ⱥ����</STRONG></TD>
      <TD width="1471" class=innerbiaoti></TD>
    </TR>
    <TR class=list>
      <TD colspan="2">
        <label>
          <select name="sendname" id="sendname" multiple  class="select" style="width:80%;height:200px">
          <%
	dim rssendname
	set rssendname =conn.Exec ("select * from {prefix}smsnum","r1")
	
	Do While Not rssendname.Eof			
		echo "<option value="""& rssendname(0) &""">"&rssendname(1) &"("&rssendname(2)&")</option>" & vbcr
		rssendname.MoveNext
	Loop
	rssendname.Close	:	Set rssendname=Nothing
		  %>
          </select>
          </label>
            </TD>
      </TR>
  <TR class=list>
      <TD colspan="2"><textarea class="textarea" name="sendmsg" cols="60" style="width:80%" rows="8"></textarea></TD>
      </TR>
    <TR class=list>
      <TD colspan="2"><INPUT class="button" type="button" value="���Ƿ�����" onClick="alert('����������')"/> <INPUT class="button" type="submit" value="����" onClick="if(confirm('ȷ������Ҫ������')){form.action='?action=sendname';}else{return false};"/></TD>
      </TR>
  </TBODY>
</TABLE></form>
</DIV>
<DIV class=tablezone id=hsendsort><form name="sendsort" method="post" action="">
  <TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
  <TBODY>
    <TR>
      <TD width="288" class=innerbiaoti><STRONG>���ѡȡȺ�����Ⱥ����</STRONG></TD>
      <TD width="1471" class=innerbiaoti></TD>
    </TR>
    <TR class=list>
      <TD colspan="2">
        
          <%
	set rssendname =conn.Exec ("select * from {prefix}smssort","r1")
	
	Do While Not rssendname.Eof			
		echo "<input type=""checkbox"" name=""sendsort"" id=""sendsort"" value="""& rssendname(0) &""">"&rssendname(1) &"</option>" & vbcr
		rssendname.MoveNext
	Loop
	rssendname.Close	:	Set rssendname=Nothing
		  %>
        
            <label>
            
            </label></TD>
      </TR>
    <TR class=list>
      <TD colspan="2"><textarea class="textarea" name="sendmsg" cols="60" style="width:80%" rows="8"></textarea></TD>
      </TR>
    <TR class=list>
      <TD colspan="2"><INPUT class="button" type="button" value="���Ƿ�����" onClick="alert('����������')"/> <INPUT class="button" type="submit" value="����" onClick="if(confirm('ȷ������Ҫ������')){form.action='?action=sendsort';}else{return false};"/></TD>
      </TR>
  </TBODY>
</TABLE></form>
</DIV>

<DIV class=searchzone id='balance'  >
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
   
    <TD width="80%"> 
    	<strong>ʣ��ɷ���������</strong><span style="color:#F00"><%=getBalance() %></span>&nbsp;&nbsp;&nbsp;
        <span><a href="AspCms_smsChargUp.asp" target="_self">���ϳ�ֵ</span>
    </TD>
  
  </TR></TBODY></TABLE>
</DIV>


<DIV class=searchzone  >
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
   
    <TD width="80%"> <strong>���ͼ�¼</strong></TD>
  
  </TR></TBODY></TABLE>
</DIV>
<FORM name="" action="" method="post">
<DIV class=listzone id=hsendrecord >
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 id="listtable">
  <TBODY>
  <TR class=list>
    <TD width=47 align="center" class=biaoti>���</TD>
    <TD width=305 align="center" class=biaoti>����ʱ��</TD>
    <TD width="173" height=28 align="center" class=biaoti><span class="searchzone">��������</span></TD>
    <TD width="207" align="center" class=biaoti>�Ƿ�ɹ�</TD>
    <TD width=1117 align="center" class=biaoti><span class="searchzone">��������</span></TD>
    </TR>
	<%smssendList()%>
    </TBODY></TABLE>
</DIV>
</FORM><br><br><br>
</BODY></HTML>
