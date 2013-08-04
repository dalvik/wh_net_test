<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="inc/aliyun_Config.asp" -->
<!--#include file="inc/aliyun_Funcs.asp" -->
<%
CheckLogin()
dim action : action=getForm("action","get")
Select case action	
	case "save" : saveconfig
End Select

Sub saveconfig
	dim configStr,configPath
	configPath="inc/aliyun_Config.asp"		
	configStr=loadFile(configPath)
	configStr = AddParas("OSS_ACCESS_ID","(\S*?)",configStr,false)	
	configStr = AddParas("OSS_ACCESS_KEY","(\S*?)",configStr,false)	
	'configStr = AddParas("DEFAULT_OSS_HOST","(\S*?)",configStr,false)
	'configStr = AddParas("BUCKET_PREFIX","(\S*?)",configStr,false)
	'configStr = AddParas("DEFAULT_ACL","(\S*?)",configStr,false)
	'configStr = AddParas("OSS_CONTENT_MD5","(\S*?)",configStr,false)
	createTextFile configStr,configPath,""
	alertMsgAndGo "修改成功",getPageName()
End Sub

Function AddParas(p,t,c,e)
	dim endTag,isString,val
	isString = ( t = "(\S*?)" )	
	val = getForm(p,"post")
	endTag = "%" & ">"
	dim templateobj
	set templateobj=new TemplateClass
	if instr(c,"Const "&p&"") = 0 then	
		c = replace(c,endTag,"")
		c = c & vbCrLf & vbCrLf
		if isString then
			c = c & "Const "&p&"="""&val&""""
		else
			if val = "" then val = 0
			c = c & "Const "&p&"="&val&""
		end if
		c = c & vbCrLf & vbCrLf & endTag
	else	
		if isString then
			if e then
				c=templateobj.regExpReplace(c,"Const "&p&"="""&t&"""","Const "&p&"="""&encode(val)&"""")
			else			
				c=templateobj.regExpReplace(c,"Const "&p&"="""&t&"""","Const "&p&"="""&val&"""")
			end if	
		else
			if val = "" then val = 0	
			c=templateobj.regExpReplace(c,"Const "&p&"="&t,"Const "&p&"="&val)
		end if
	end if
	set templateobj = nothing
	AddParas = c
End Function
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
    <TD height=30><span onClick="location.href='aliyun_EditConfig.asp'" class="head3" >编辑配置信息</span></TD>
    <TD><span onClick="location.href='aliyun_ListBucket.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'" >Bucket列表</span> </TD>
    <TD><span onClick="location.href='aliyun_ListObject.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'">备份数据列表</span> </TD>
    <TD><span onClick="location.href='aliyun_Backup.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'">一键数据备份</span> </TD>
  </TR>
</TBODY>
</TABLE>
</DIV>

<DIV class=formzone>
<FORM name="aliyunedit" id="aliyunedit" action="?action=save" method="post">
<DIV class=tablezone>
<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
<TBODY>

    <TR>
    	<TD width="193" class=innerbiaoti><STRONG>云备份配置--基本信息</STRONG></TD>
        <TD width="568" class=innerbiaoti></TD>
    </TR>
  	<TR class=list>
        <TD>阿里云开放平台ACCESS_ID: </TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;"  value="<%=OSS_ACCESS_ID%>" name="OSS_ACCESS_ID"/>&nbsp;&nbsp;<font color="#FF0000">*</font></TD>
    </TR>
    <TR class=list>
        <TD>阿里云开放平台OSS_ACCESS_KEY: </TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;"  value="<%=OSS_ACCESS_KEY%>" name="OSS_ACCESS_KEY"/>&nbsp;&nbsp;<font color="#FF0000">*</font></TD>
    </TR>
    <TR class=list>
        <TD colspan="2"><a href="http://oss.aliyun.com/guide/details?id=257" target="_blank">如何获取ID和KEY？</a> </TD>
    </TR>
    <TR>
    	<TD width="193" class=innerbiaoti><STRONG>云备份配置--高级信息</STRONG></TD>
        <TD width="568" class=innerbiaoti></TD>
    </TR>
    <TR class=list>
        <TD>阿里云OSS服务公网地址: </TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;background:#C9C9C9;"  value="<%=DEFAULT_OSS_HOST%>" name="DEFAULT_OSS_HOST" disabled/></TD>
    </TR>
    <TR class=list>
        <TD>Bucket前缀<Br>(不能为空，且只能是英文字母或数字，不能超过10个字符): </TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;background:#C9C9C9;"  value="<%=BUCKET_PREFIX%>" name="BUCKET_PREFIX" disabled /></TD>
    </TR>
    <TR class=list>
        <TD>默认创建的bucket的权限: </TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;background:#C9C9C9;"  value="<%=DEFAULT_ACL%>" name="DEFAULT_ACL" disabled /></TD>
    </TR>
    <TR class=list>
        <TD>header头部MD5加密码: </TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;background:#C9C9C9;"  value="<%=OSS_CONTENT_MD5%>" name="OSS_CONTENT_MD5" disabled /></TD>
    </TR>
</TBODY></TABLE>

</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="保存信息" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 
 </DIV></FORM>
</DIV>
</BODY>
</HTML>