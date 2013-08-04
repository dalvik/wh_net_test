<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="inc/aliyun_Config.asp" -->
<!--#include file="inc/aliyun_Funcs.asp" -->
<%
CheckLogin()
dim bucketname : bucketname=getForm("name","get")
dim action : action=getForm("action","get")
dim aclnew :aclnew=getForm("aclNew","post")
if bucketname="" then
	alertMsgAndGo "bucket名称不能为空","aliyun_ListBucket.asp"
end if

if action="editacl" then
	changeAcl
end if

Dim dAcl
Set dAcl=Server.CreateObject("Scripting.Dictionary")
dAcl.Add "private","私有权限"
dAcl.Add "public-read","公开只读权限"
dAcl.Add "public-read-write","公开读写权限"

Sub changeAcl
	Dim rt:rt=set_bucket_acl(bucketname,aclnew)
	if rt="1" then
  		alertMsgAndGo "更改成功","aliyun_EditAcl.asp?name="&bucketname
 	else
 		alertMsgAndGo rt,"aliyun_EditAcl.asp?name="&bucketname
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
    <TD height=30><span onClick="location.href='aliyun_EditConfig.asp'" class="head2" onmouseover="this.className='head4'" onmouseout="this.className='head2'">编辑配置信息</span></TD>
    <TD><span onClick="location.href='aliyun_ListBucket.asp'" class="head3" >Bucket列表</span> </TD>
    <TD><span onClick="location.href='aliyun_ListObject.asp'" class="head2" onmouseover="this.className='head4'" onmouseout="this.className='head2'">备份数据列表</span> </TD>
    <TD><span onClick="location.href='aliyun_Backup.asp'" class="head2" onmouseover="this.className='head4'" onmouseout="this.className='head2'">一键数据备份</span> </TD>
  </TR>
</TBODY>
</TABLE>
</DIV>

<DIV class=formzone>
<FORM name="aliyunedit" id="aliyunedit" action="?action=editacl&name=<%=bucketname%>" method="post">
<DIV class=tablezone>
<TABLE cellSpacing=0 cellPadding=8 width="60%" align=center border=0>
<TBODY>
	
    <TR>
    	<TD width="193" class=innerbiaoti><STRONG><%=bucketname%>ACL权限修改</STRONG></TD>
        <TD width="568" class=innerbiaoti></TD>
    </TR>
  	<TR class=list>
        <TD><%=bucketname%>原有ACL权限: </TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;"  value="<%=dAcl.item(get_bucket_acl(bucketname))%>" name="aclOld" disabled/></TD>
    </TR>
    <TR class=list>
        <TD>选择ACL权限更改: </TD>
        <TD align="left">
        	<select name="aclNew">
            	<option value="private" selected >私有权限</option>
                <option value="public-read">公开只读权限</option>
                <option value="public-read-write">公开读写权限</option>
            </select>
        </TD>
    </TR>
</TBODY></TABLE>

</DIV>
<DIV class=adminsubmit style="margin-left:35%;">
<INPUT class="button" type="submit" value="提交并更改信息" />&nbsp;&nbsp;&nbsp;&nbsp;
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 
 </DIV></FORM>
</DIV>

<DIV class=searchzone>
<DIV class=tablezone>
<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
<TBODY>

    <TR>
    	<TD width="193" class=innerbiaoti><STRONG>关于ACL权限的说明</STRONG></TD>
        <TD width="568" class=innerbiaoti></TD>
    </TR>
  	<TR class=list>
        <TD>private(私有权限): </TD>
        <TD align="left"><p style="text-indent:20px">只有该Bucket的创建者可以对该Bucket内的Object进行读写操作（包括Put、Delete和Get Object）；其他人无法访问该Bucket内的Object。</p>
        </TD>
    </TR>
    <TR class=list>
        <TD>public-read(公开只读权限): </TD>
        <TD align="left"><p style="text-indent:20px">只有该Bucket的创建者可以对该Bucket内的Object进行写操作（包括PUT和Delete Object）；任何人（包括匿名访问）可以对该Bucket中的Object进行读操作（Get Object）</p>
        </TD>
    </TR>
    <TR class=list>
        <TD>public-read-write(公开读写权限): </TD>
        <TD align="left"><p style="text-indent:20px">任何人（包括匿名访问）都可以对该Bucket中的Object进行PUT，Get和Delete操作；所有这些操作产生的费用由该Bucket的创建者承担，请慎用该权限</p>
        </TD>
    </TR>
</TBODY></TABLE>
</DIV>
</DIV>


</BODY>
</HTML>