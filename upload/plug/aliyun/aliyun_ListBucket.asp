<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="inc/aliyun_Config.asp" -->
<!--#include file="inc/aliyun_Funcs.asp" -->
<%
CheckLogin()

Dim dAcl
Set dAcl=Server.CreateObject("Scripting.Dictionary")
dAcl.Add "private","˽��Ȩ��"
dAcl.Add "public-read","����ֻ��Ȩ��"
dAcl.Add "public-read-write","������дȨ��"


Sub bucketlist
	Dim blist,i
	blist=bucket_list()
	for i=0 to UBound(blist)-1
		echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf& _
				"<td>"&blist(i)&"</td>"&vbcrlf& _
				"<td>"&dAcl.item(get_bucket_acl(blist(i)))&"</td>"&vbcrlf& _
				"<td>"&"<span class=""button2""><a href=""aliyun_EditAcl.asp?name="&blist(i)&""" >����ACLȨ��</a></span>&nbsp;&nbsp;"&vbcrlf& _
					"<span class=""button2""><a href=""aliyun_ListObject.asp?year="&split(blist(i),"_")(1)&""">�鿴�����ļ�</a></span> "&"</td>"&vbcrlf& _
				"</tr>"&vbcrlf
	next
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<LINK href="images/style.css" type=text/css rel=stylesheet>
</HEAD>
<BODY>
<DIV class=searchzone>
<style>
.head2 {
	BORDER-RIGHT: #cde6ff 1px solid; PADDING-RIGHT: 20px; BORDER-TOP: #cde6ff 1px solid; PADDING-LEFT: 20px; BACKGROUND: #e4f2ff; PADDING-BOTTOM: 4px; FONT: 12px/20px Verdana, Arial, Helvetica, sans-serif; BORDER-LEFT: #cde6ff 1px solid; COLOR: #555; PADDING-TOP: 4px; BORDER-BOTTOM: #cde6ff 1px solid; cursor: pointer;
}
.head3 {
	BORDER-RIGHT: #cde6ff 1px solid; PADDING-RIGHT: 20px; BORDER-TOP: #cde6ff 1px solid; PADDING-LEFT: 20px; BACKGROUND:#1A96CC; PADDING-BOTTOM: 4px; FONT: 12px/20px Verdana, Arial, Helvetica, sans-serif; BORDER-LEFT: #cde6ff 1px solid; COLOR: #fff; PADDING-TOP: 4px; BORDER-BOTTOM: #cde6ff 1px solid; cursor: pointer;
}
.head4 {
	BORDER-RIGHT: #cde6ff 1px solid; PADDING-RIGHT: 20px; BORDER-TOP: #cde6ff 1px solid; PADDING-LEFT: 20px; BACKGROUND: #B6E2F5; PADDING-BOTTOM: 4px; FONT: 12px/20px Verdana, Arial, Helvetica, sans-serif; BORDER-LEFT: #cde6ff 1px solid; COLOR: #fff; PADDING-TOP: 4px; BORDER-BOTTOM: #cde6ff 1px solid; cursor: pointer;
}
</style>
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
  <TR align="center">
    <TD height=30><span onClick="location.href='aliyun_EditConfig.asp'" class="head2" onmouseover="this.className='head4'" onmouseout="this.className='head2'">�༭������Ϣ</span></TD>
    <TD><span onClick="location.href='aliyun_ListBucket.asp'" class="head3" >Bucket�б�</span> </TD>
    <TD><span onClick="location.href='aliyun_ListObject.asp'" class="head2" onmouseover="this.className='head4'" onmouseout="this.className='head2'">���������б�</span> </TD>
    <TD><span onClick="location.href='aliyun_Backup.asp'" class="head2" onmouseover="this.className='head4'" onmouseout="this.className='head2'">һ�����ݱ���</span> </TD>
  </TR>
</TBODY>
</TABLE>
</DIV>


<DIV class=listzone>
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border="1" bordercolor="#cde6ff" >
  <TBODY>
  <TR class=list>
    <TD width=100 align="center" class=biaoti>Bucket����</TD>
    <TD width=150 align="center" class=biaoti>Bucket ACL</TD>
    <TD height=28 align="center" class=biaoti>����</TD>
    </TR>
	<%bucketlist%>
    </TBODY></TABLE>
</DIV>
<DIV class=searchzone><TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0><TBODY><TR height="30"><td>&nbsp;</td></TR></TBODY></TABLE></DIV>

</BODY>
</HTML>