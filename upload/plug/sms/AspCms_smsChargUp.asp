<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="inc/AspCms_smsFuncs.asp" -->
<!--#include file="inc/AspCms_smsConfig.asp" -->
<%CheckLogin()%>
<%
dim action : action=getForm("action","get")
dim msg,cardno,cardpwd
if smsIsActivate=1 then 
	msg="剩余可发送<span style='color:#F00;font-size:12px;'>"&getBalance()&"</span>条"
else
	msg="您尚未激活序列号，<span style='color:#f00'>点击激活</span>"
end if

Select case action	
	case "chargup" :smschargeup 
End Select

Sub smschargeup
	dim returnmsg
	returnmsg=split(ChargUp(cardno,cardpwd)," ")
	if returnmsg(0)="0" then 
		alertMsgAndGo "充值成功",getPageName()
	else
		alertMsgAndGo "充值失败,"&returnmsg(1),getPageName()
	end if	
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<LINK href="images/style.css" type=text/css rel=stylesheet>
</HEAD>
<BODY>
<!--#include file="AspCms_smsHead.asp" -->
<DIV class=formzone>
<FORM name="smsreg" id="smsreg" action="?action=chargup" method="post">
<DIV class=tablezone>
<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
<TBODY>

    <TR>
    	<TD width="193" class=innerbiaoti>
        	<STRONG>充值卡充值</STRONG>&nbsp;(<%=msg %>)</TD>
        <TD width="568" class=innerbiaoti></TD>
    </TR>
    <TR class=list>
        <TD>充值卡号（由商务处获得） : </TD>
        <TD><input class="input" size="30" style="width:300px;" name="cardno"/></TD>
    </TR>
    <TR class=list>
        <TD>充值卡密码（由商务处获得） : </TD>
        <TD><input class="input" size="30" style="width:300px;" name="cardpwd"/></TD>
    </TR>
</TBODY></TABLE>
<br>
<br>
<br>
<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
  <TBODY>
    <TR>
      <TD width="288" class=innerbiaoti><STRONG>优惠充值表</STRONG></TD>
      <TD width="1471" class=innerbiaoti></TD>
    </TR>
    <TR class=list>
      <TD>短信</TD>
      <TD align="left">&nbsp;</TD>
    </TR>
    <TR class=list>
      <TD>1000条</TD>
      <TD>100元</TD>
    </TR>
    <TR class=list>
      <TD>2000</TD>
      <TD>190元</TD>
    </TR>
    <TR class=list>
      <TD>5000</TD>
      <TD>450元</TD>
    </TR>
    <TR class=list>
      <TD>彩信</TD>
      <TD>&nbsp;</TD>
    </TR>
    <TR class=list>
      <TD height="30">&nbsp;</TD>
      <TD>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="现在充值" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 
 </DIV></FORM>
</DIV>
</BODY>
</HTML>
