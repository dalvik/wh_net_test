<!--#include file="AspCms_SettingFun.asp" -->
<%CheckAdmin("AspCms_PaymentSetting.asp")%>
<%getComplanySetting%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<LINK href="../images/style.css" type=text/css rel=stylesheet>
</HEAD>
<BODY>
<DIV class=formzone>
<FORM name="" action="?action=saveb" method="post">
<DIV class=tablezone>

<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
<TBODY>
<TR>
<TD width="193" class=innerbiaoti>在线支付</TD>
<TD class=innerbiaoti width=568 height=28></TD>
</TR>
<TR class=list>
<TD>帐号信息 : </TD>
<TD><textarea class="textarea" name="Payment_Online" cols="60" style="width:98%" rows="8"><%SafeEcho("Payment_Online")%></textarea></TD>
</TR>

<TR>
<TD width="193" class=innerbiaoti>银行支付</TD>
<TD class=innerbiaoti width=568 height=28></TD>
</TR>
<TR class=list>
<TD>帐号信息 : </TD>
<TD><textarea class="textarea" name="Payment_Bank" cols="60" style="width:98%" rows="8"><%SafeEcho("Payment_Bank")%></textarea></TD>
</TR>

<TR>
<TD width="193" class=innerbiaoti>邮局付款</TD>
<TD class=innerbiaoti width=568 height=28></TD>
</TR>
<TR class=list>
<TD>帐号信息 : </TD>
<TD><textarea class="textarea" name="Payment_PostOffice" cols="60" style="width:98%" rows="8"><%SafeEcho("Payment_PostOffice")%></textarea></TD>
</TR>

    
</TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT type="hidden" value="<%=LanguageID%>" name="LanguageID" />
<INPUT class="button" type="submit" value="保存" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 
 </DIV></FORM>
</DIV>
</BODY>
</HTML>
