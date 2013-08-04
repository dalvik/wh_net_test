<!--#include file="AspCms_CustomFun.asp" -->
<%CheckAdmin("AspCms_Custom.asp")%>
<%getContent%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>
</HEAD>
<BODY>
<FORM name="form" action="?action=edit&page=<%=page%>&order=<%=order%>&sort=<%=sortID%>&keyword=<%=keyword%>" method="post" >
<DIV class=formzone>
<DIV class=namezone>查看表单信息</DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
<TBODY>
<TR>						
    <TD align=middle width=100 height=30>编号</TD>
    <TD><%=Customid%></TD>
</TR>
<TR>
    <TD align=middle width=100 height=30>时间</TD>
    <TD><%=addtime%></TD>
</TR>

<%getaForm(ID)%>

  
</TBODY>
</TABLE>
</DIV>
<DIV class=adminsubmit></DIV>
</DIV></FORM>

</BODY></HTML>    

            
