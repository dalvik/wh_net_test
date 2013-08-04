<!--#include file="AspCms_StyleFun.asp" -->
<%CheckAdmin("AspCms_Template.asp")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../images/style.css" type=text/css rel=stylesheet>
</HEAD>
<BODY>
	  <form action="?acttype=<%=acttype%>&action=add" method="post" name="form">
<DIV class=formzone>
<DIV class=namezone添加<%=nametype%></DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>						
    <TD align=middle width=100 height=30>文件名称</TD>
    <TD><input class="input" type="text" name="filename" /> 
    <FONT color=#ff0000>*</FONT> (允许文件名有<%=Template_AllowFileExt%>)</TD>
  </TR>
  <TR>
    <TD align=middle width=100 height=30>文件内容</TD>    
    <TD> <textarea class="textarea" cols="80" rows="3" name="filetext" style=" height:400px; width:98%;" ></textarea></TD>
</tr> 
  
    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="添加" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/>  
</DIV></DIV></FORM>
</BODY></HTML>