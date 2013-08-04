<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="inc/aliyun_Config.asp" -->
<!--#include file="inc/aliyun_Funcs.asp" -->
<%
CheckLogin()
dim ossYear : ossYear=getForm("year","get")
dim ossMonth : ossMonth=getForm("month","get")
dim action:action=getForm("action","get")

if action="delete" then
	objectDel
end if

if ossYear="" and ossMonth="" then 
	ossYear=year(Now)
	ossMonth=month(Now)
elseif ossYear<>"" and ossMonth="" then
	ossYear=int(ossYear)
	if ossYear=year(Now) then
		ossMonth=month(Now)
	elseif ossYear<year(Now) then
		ossMonth=12
	else
		alertMsgAndGo "年份不合法",getPageName()
	end if
elseif ossYear="" and ossMonth<>"" then
	alertMsgAndGo "参数错误",getPageName()
end if

Sub objectDel
	Dim object,bucket,rtxt,arr
	object=getForm("name","get")
	arr=split(object,"_")
	bucket=BUCKET_PREFIX&"_"&arr(0)
	object=object&""&OSS_OBJECT_EXT
	rtxt=object_delete(bucket,object)
	if rtxt="1" then
		alertMsgAndGo "成功删除文件"&object,Request.ServerVariables("Script_Name")&"?year="&arr(0)&"&month="&arr(1)
	else
		alertMsgAndGo "删除失败，原因："&rtxt,Request.ServerVariables("Script_Name")&"?year="&arr(0)&"&month="&arr(1)
	end if
End Sub

Sub listYear
	Dim startYear,endYear,iter
	startYear=2011
	endYear=year(now)
	iter=startYear
	while iter<=endYear
		if iter=int(ossYear) then
			echo "<option value="&iter&" selected >"&iter&"</option>"
		else
			echo "<option value="&iter&" >"&iter&"</option>"
		end if
		iter=iter+1
	wend
End Sub

Sub listMonth
	Dim iter
	for iter=1 to 12
		if iter=int(ossMonth) then
			echo "<option value="&iter&" selected >"&iter&"</option>"
		else
			echo "<option value="&iter&" >"&iter&"</option>"
		end if
	next
End Sub

Sub objectlist
	Dim objectarr,objecti,file_name,file_size,file_time,file_url
	objectarr=object_list(BUCKET_PREFIX&"_"&ossYear,ossYear&"_"&ossMonth&"_",OSS_OBJECT_MAXLEN)
	if not isArray(objectarr) then
		alertMsgAndGo "参数错误",getPageName()
	end if
	if UBound(objectarr)=0 then
		echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf& _
				"<td colspan=""4"" style=""color:#f00"">无返回值</td>"&vbcrlf& _
				"</tr>"&vbcrlf
	else
		for objecti=0 to UBound(objectarr)-1
			file_name=objectarr(objecti,0)
			file_time=split(objectarr(objecti,1),"T")(0)
			file_size=int(objectarr(objecti,2)/1000)
			if file_size>1000 then
				file_size=FormatNumber(file_size/1000,1)&"M"
			else
				file_size=file_size&"K"
			end if
			'生成下载地址
			file_url=get_sign_url(BUCKET_PREFIX&"_"&ossYear,file_name)
			echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf& _
					"<td><a href="""&file_url&""">"&file_name&"</a></td>"&vbcrlf& _
					"<td>"&file_size&"</td>"&vbcrlf& _
					"<td>"&file_time&"</td>"&vbcrlf& _
					"<td><span><a href="""&file_url&""" class='button'>下载</a></span>&nbsp;&nbsp;|&nbsp;&nbsp;<span><a onclick=""DelObject('delete','"&split(file_name,".")(0)&"')"" href=# class='button'>删除</a></span></td>"&vbcrlf& _
					"</tr>"&vbcrlf
			
		next
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
<script>
function DelObject(action,name){
    if(confirm('确认删除已选择数据吗?')){
        window.location.href=window.location.pathname+"?action="+action+"&name="+name;
		//alert(window.location.pathname+"?action="+action+"&name="+name);
    }
}
</script>
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
    <TD><span onClick="location.href='aliyun_ListBucket.asp'" class="head2" onmouseover="this.className='head4'" onmouseout="this.className='head2'" >Bucket列表</span> </TD>
    <TD><span onClick="location.href='aliyun_ListObject.asp'" class="head3" >备份数据列表</span> </TD>
    <TD><span onClick="location.href='aliyun_Backup.asp'" class="head2" onmouseover="this.className='head4'" onmouseout="this.className='head2'">一键数据备份</span> </TD>
  </TR>
</TBODY>
</TABLE>
</DIV>

<DIV class=searchzone style="height:30px;">
<form action="" method="get" >
<style>
.button2 {
	BORDER-RIGHT: #cde6ff 1px solid; PADDING-RIGHT: 20px; BORDER-TOP: #cde6ff 1px solid; PADDING-LEFT: 20px; BACKGROUND: #e4f2ff; PADDING-BOTTOM: 4px; FONT: 12px/20px Verdana, Arial, Helvetica, sans-serif; BORDER-LEFT: #cde6ff 1px solid; COLOR: #555; PADDING-TOP: 4px; BORDER-BOTTOM: #cde6ff 1px solid; cursor: pointer;
}
</style>
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" align="left" border=0>
<TBODY>
  <TR align="center">
    <TD height=30 width="150px">
    	<label>年份：</label>
        <select name="year" class="button">
        	<%listYear%>
        </select>
    </TD>
    <TD width="150px">
    	<label>月份：</label>
        <select name="month" class="button">
        	<%listMonth%>
        </select>
    </TD>
    <TD align='left'><INPUT class="button" type="submit" value="查找" style="padding-left:15px;padding-right:15px;" /></TD>
  </TR>
</TBODY>
</TABLE>
</form>
</DIV>

<DIV class=listzone>
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border="1" bordercolor="#cde6ff" >
  <TBODY>
  <TR class=list>
    <TD width=25% align="center" class=biaoti>文件名</TD>
    <TD width=15% align="center" class=biaoti>文件大小</TD>
    <TD width=25% align="center" class=biaoti>上传时间</TD>
    <TD height=28 align="center" class=biaoti>操作</TD>
    </TR>
	<%objectlist%>
    </TBODY></TABLE>
</DIV>
<DIV class=searchzone><TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0><TBODY><TR height="30"><td>&nbsp;</td></TR></TBODY></TABLE></DIV>
</BODY>
</HTML>