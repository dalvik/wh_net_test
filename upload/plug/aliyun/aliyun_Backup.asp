<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="inc/aliyun_Config.asp" -->
<!--#include file="inc/aliyun_Funcs.asp" -->
<%
CheckLogin()

Dim action:action=getForm("action","get")
if action="backup" then
	startBackup
end if


Dim bakdate,bakfile
bakdate=year(now)&"-"&month(now)&"-"&day(now)
bakfile=year(now)&"_"&month(now)&"_"&day(now)


Sub startBackup
	Server.ScriptTimeOut=1000
	Dim bucket,object,cbak,rbak
	Dim datapath:datapath=sitePath&"/"&accessFilePath   '备份文件位置
	object=getForm("bakfile","post")
	object=object&""&OSS_OBJECT_EXT
	bucket=BUCKET_PREFIX&"_"&year(now)
	cbak=bucket_create(bucket,DEFAULT_ACL)   '若bucket不存在时，创建bucket
	if cbak<>"1" then 
		alertMsgAndGo cbak,getPageName()  '创建失败
	end if
	response.Write("<br><br><br><center>正在备份中，请耐心等待。。。。</center>")
echo datapath
	rbak=upload_file_by_content(bucket,object,datapath)
	if rbak<>"1" then 
		alertMsgAndGo rbak,getPageName()  '上传失败
	end if
	'更新上次备份时间为当前时间
	'response.Write(rbak)
	dim configStr,configPath
	configPath="inc/aliyun_Config.asp"		
	configStr=loadFile(configPath)
	configStr = AddParas("PAST_BACKUP","(\S*?)",configStr,false)	
	createTextFile configStr,configPath,""
	alertMsgAndGo "备份成功",getPageName()
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
    <TD height=30><span onClick="location.href='aliyun_EditConfig.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'">编辑配置信息</span></TD>
    <TD><span onClick="location.href='aliyun_ListBucket.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'" >Bucket列表</span> </TD>
    <TD><span onClick="location.href='aliyun_ListObject.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'">备份数据列表</span> </TD>
    <TD><span onClick="location.href='aliyun_Backup.asp'" class="head3" >一键数据备份</span> </TD>
  </TR>
</TBODY>
</TABLE>
</DIV>
<script>
function doconfirm()
{
	 if(confirm('备份数据时会覆盖同名备份文件，确认要立刻备份吗?')){
		 alert("文件备份中，请耐心等待，不要重复点击备份");
		 document.getElementById('aliyunedit').submit();
	 }
	 
}
</script>
<DIV class=formzone>
<FORM name="aliyunedit" id="aliyunedit" action="?action=backup" method="post">
<input type="hidden" name="PAST_BACKUP" value="<%=bakdate%>" >
<input type="hidden" name="bakfile" value="<%=bakfile%>" >
<DIV class=tablezone>
<TABLE cellSpacing=0 cellPadding=8 width="70%" align=center border=1 bordercolor=" #cde6ff">
<TBODY>

    <TR>
    	<TD width="60%" class=innerbiaoti height="50px" style="font-size:36px;"><img src="images/aliyun.gif"></TD>
        <TD width="40%" class=innerbiaoti></TD>
    </TR>
  	<TR class=list>
        <TD align="left" height="300px"><p style="text-indent:20px;margin:10px;font-size:14px;line-height:35px;">阿里云存储服务（Open Storage Service，简称OSS），是阿里云对外提供的海量，安全，低成本，高可靠的云存储服务。用户可以通过简单的REST接口，在任何时间、任何地点上传和下载数据，也可以使用WEB页面对数据进行管理。基于OSS，用户可以搭建出各种多媒体分享网站、网盘、个人企业数据备份等基于大规模数据的服务。</p><br>
        <span style="margin-left:30px;margin-top:10px;color:#F00">温馨提示一：文件过大时，备份时间也会变长，请耐心等待。。。</span><br><br>
        <span style="margin-left:30px;color:#F00">温馨提示二：待备份文件必须为可读权限，否则无法读取数据。。。</span>
        </TD>
        <TD align="left" style="font-size:14px;padding-left:45px;">
        	<span>待备份文件路径：<font color="#FF0000"><%=sitePath&"/"&accessFilePath%></font></span><br><br>
        	<span>上次备份时间：<font color="#FF0000"><%=PAST_BACKUP%></font></span><br><br>
        	<span>当前备份时间：<font color="#FF0000"><%=bakdate%></font></span><br><br>
            <span>将要备份生成的文件名：<font color="#FF0000"><%=bakfile&""&OSS_OBJECT_EXT%></font></span><br><br>
            <!--<span><input type="image" src="images/bakup.gif" width="200"></span>-->
            <span><img src="images/bakup.gif" width="200" onClick="doconfirm();" style="cursor:pointer;"></span>
        </TD>
    </TR>
</TBODY></TABLE>

</DIV>
</FORM>
</DIV>

</BODY>
</HTML>