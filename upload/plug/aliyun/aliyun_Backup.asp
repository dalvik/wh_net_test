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
	Dim datapath:datapath=sitePath&"/"&accessFilePath   '�����ļ�λ��
	object=getForm("bakfile","post")
	object=object&""&OSS_OBJECT_EXT
	bucket=BUCKET_PREFIX&"_"&year(now)
	cbak=bucket_create(bucket,DEFAULT_ACL)   '��bucket������ʱ������bucket
	if cbak<>"1" then 
		alertMsgAndGo cbak,getPageName()  '����ʧ��
	end if
	response.Write("<br><br><br><center>���ڱ����У������ĵȴ���������</center>")
echo datapath
	rbak=upload_file_by_content(bucket,object,datapath)
	if rbak<>"1" then 
		alertMsgAndGo rbak,getPageName()  '�ϴ�ʧ��
	end if
	'�����ϴα���ʱ��Ϊ��ǰʱ��
	'response.Write(rbak)
	dim configStr,configPath
	configPath="inc/aliyun_Config.asp"		
	configStr=loadFile(configPath)
	configStr = AddParas("PAST_BACKUP","(\S*?)",configStr,false)	
	createTextFile configStr,configPath,""
	alertMsgAndGo "���ݳɹ�",getPageName()
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
    <TD height=30><span onClick="location.href='aliyun_EditConfig.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'">�༭������Ϣ</span></TD>
    <TD><span onClick="location.href='aliyun_ListBucket.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'" >Bucket�б�</span> </TD>
    <TD><span onClick="location.href='aliyun_ListObject.asp'" class="head2" onMouseOver="this.className='head4'" onMouseOut="this.className='head2'">���������б�</span> </TD>
    <TD><span onClick="location.href='aliyun_Backup.asp'" class="head3" >һ�����ݱ���</span> </TD>
  </TR>
</TBODY>
</TABLE>
</DIV>
<script>
function doconfirm()
{
	 if(confirm('��������ʱ�Ḳ��ͬ�������ļ���ȷ��Ҫ���̱�����?')){
		 alert("�ļ������У������ĵȴ�����Ҫ�ظ��������");
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
        <TD align="left" height="300px"><p style="text-indent:20px;margin:10px;font-size:14px;line-height:35px;">�����ƴ洢����Open Storage Service�����OSS�����ǰ����ƶ����ṩ�ĺ�������ȫ���ͳɱ����߿ɿ����ƴ洢�����û�����ͨ���򵥵�REST�ӿڣ����κ�ʱ�䡢�κεص��ϴ����������ݣ�Ҳ����ʹ��WEBҳ������ݽ��й�������OSS���û����Դ�����ֶ�ý�������վ�����̡�������ҵ���ݱ��ݵȻ��ڴ��ģ���ݵķ���</p><br>
        <span style="margin-left:30px;margin-top:10px;color:#F00">��ܰ��ʾһ���ļ�����ʱ������ʱ��Ҳ��䳤�������ĵȴ�������</span><br><br>
        <span style="margin-left:30px;color:#F00">��ܰ��ʾ�����������ļ�����Ϊ�ɶ�Ȩ�ޣ������޷���ȡ���ݡ�����</span>
        </TD>
        <TD align="left" style="font-size:14px;padding-left:45px;">
        	<span>�������ļ�·����<font color="#FF0000"><%=sitePath&"/"&accessFilePath%></font></span><br><br>
        	<span>�ϴα���ʱ�䣺<font color="#FF0000"><%=PAST_BACKUP%></font></span><br><br>
        	<span>��ǰ����ʱ�䣺<font color="#FF0000"><%=bakdate%></font></span><br><br>
            <span>��Ҫ�������ɵ��ļ�����<font color="#FF0000"><%=bakfile&""&OSS_OBJECT_EXT%></font></span><br><br>
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