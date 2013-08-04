<!--#include file="AspCms_smsSortFun.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="images/style.css" type=text/css rel=stylesheet>

</HEAD>
<BODY>
<!--#include file="AspCms_smsHead.asp" -->
<FORM name="form" action="?action=addnum" method="post" >
<DIV class=formzone>
<DIV class=namezone>添加号码</DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>						
    <TD align=middle width=100 height=30>姓名</TD>
    <TD><INPUT class="input" style="WIDTH: 200px" maxLength="200" name="smsname"/> </TD>
  </TR>
  <TR id="hid9">
    <TD align=middle width=100 height=30>电话</TD>
    <TD><input class="input" maxlength="200" style="WIDTH: 200px" name="smsnum"/></TD></TR>
  <TR>
    <TD align=middle height=30>电话分类</TD>
    <TD><%loadsmsSelect 0%></TD>
  </TR>
  <TR>
    <TD align=middle height=30>备注</TD>
    <TD>  <TEXTAREA class="textarea" style="WIDTH: 450px" name="remark" rows="3"></TEXTAREA>    </TD>
  </TR>
    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="添加" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 

<INPUT onClick="location.href='<%=getPageName()%>'" type="button" value="刷新" class="button" /> 
</DIV></DIV></FORM>

</BODY></HTML>