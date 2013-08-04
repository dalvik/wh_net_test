<!--#include file="vote_SettingFun.asp" -->
<!--#include file="vote_config.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>投票结果</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK href="style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.3790.4807" name=GENERATOR>
</HEAD>
<BODY>
<%
if request.Cookies("vote")<>"yes" then
	echo "<script>alert('请先投票');window.close();</script>"
	die
end if
dim votenames,votenums,i,votecount
votenames=split(votename,"|")
votenums=split(votenum,"|")
votecount=0
for i=0 to ubound(votenums)
	votecount=cint(votecount)+cint(votenums(i))
next
%>

<DIV class=formzone>
  <strong>投票结果</strong>
<DIV class=tablezone><table width="100%" border="1">
  <tr>
    <td colspan="3"><%=votetitle%></td>
  </tr>
  
  <%for i=0 to ubound(votenames)%>
  <tr>
    <td width="25%"><%=votenames(i)%></td>
    <td width="65%"> <table width="<%=(votenums(i)/votecount)*100%>%" border="0" bgcolor="#6699FF">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</td>
    <td width="10%"><%=votenums(i)%>票</td>
  </tr>
  <%next%>
</table></DIV>
</DIV>
</BODY>
</HTML>