<!--#include file="AspCms_SpecFun.asp" -->
<%CheckAdmin("AspCms_Spec.asp")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>

<script type="text/javascript">
function changeSpecControlType(select){
	var trSpecOptions = document.getElementById("trSpecOptions");
	if(select.options[select.selectedIndex].value>5)
		trSpecOptions.style.display = "block";
	else
		trSpecOptions.style.display = "none";
}
function changeSpecCategoryType(select){
	var preShow = document.getElementById("preShow");
	var spShow  = document.getElementById("spShow");
	preShow.innerHTML = select.options[select.selectedIndex].value;
	if(spShow.innerHTML == "")spShow.innerHTML = "_";
	
}
</script>
</HEAD>
<BODY>
<FORM action="?action=<%if request.QueryString("action")="edit" then%>editsave<%else%>add<%end if%>" method="post" name="frmSpec">
<input type="hidden" name="SpecDiversification" value="<%=SpecDiversification%>" />
<input type="hidden" name="SpecID" value="<%=SpecID%>" />
<DIV class=formzone>
<DIV class=namezone><%if request.QueryString("action")="edit" then%>编辑参数<%else%>添加参数<%end if%></DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
<TBODY>
<TR>						
<TD align=middle width=100 height=30>参数名称</TD>
<TD><INPUT class="input" style="WIDTH: 200px" maxLength="200" name="SpecName" value="<%=SpecName%>"/> <FONT color=#ff0000>*</FONT> </TD>
</TR>
<TR>
<TD align=middle width=100 height=30>字段名称</TD>    
<TD><INPUT class="input" style="WIDTH: 200px" onKeyUp="document.getElementById('endShow').innerHTML=value" maxLength="200" name="SpecField"  value="<%=SpecField%>"/> <FONT color=#ff0000>*预览 [<span id="endShow"><%=SpecField%></span>]</FONT></TD>
</tr>
<TR>
<TD align=middle width=100 height=30>必须填写</TD>    
<TD><INPUT type="checkbox" class="input" name="SpecNotNull" <%if SpecNotNull then%>checked<%end if%> /></TD>
</tr>
<TR>
<TD align=middle width=100 height=30>控件类型</TD>    
<TD><select name="SpecControlType" onChange="changeSpecControlType(this)">
<option value="0" selected>文本</option>
<option value="1">数字</option>
<option value="2">编辑器</option>

<option value="3">附件</option>
<option value="4">日期</option>
<option value="5">颜色</option>
<option value="6">单选</option>
<option value="7">多选</option>
</select>
<script>
//frmSpec.SpecControlType.selectedIndex = <%=SpecControlType%>;
for(var i=0;i<frmSpec.SpecControlType.options.length;i++){
	if(frmSpec.SpecControlType.options[i].value == "<%=SpecControlType%>")
		frmSpec.SpecControlType.selectedIndex = i;
}
</script>
</TD>
</tr>
<TR style="display:none" id="trSpecOptions">
<TD align=middle width=100 height=30>备选内容</TD>    
<TD><textarea name="SpecOptions" cols="50" rows="5" class="input" style="width:500px;height:80px;"><%=SpecOptions%></textarea>多条请用回车区别</TD>
</tr>
<TR>
<td align=middle width=100 height=30>应用与分类</td>    
<td><select name="SpecCategory" onChange="changeSpecCategoryType(this)">
<option value="P" selected>产品</option>
<option value="S">单篇</option>
<option value="C">文章</option>
<option value="DL">下载</option>
<option value="HR">招聘</option>
<option value="FO">相册</option>
<option value="VI">视频</option>
</select>
<script>
for(var i=0;i<frmSpec.SpecCategory.options.length;i++){
	if(frmSpec.SpecCategory.options[i].value == "<%=SpecCategory%>")
		frmSpec.SpecCategory.selectedIndex = i;
}
document.getElementById("preShow").innerHTML = frmSpec.SpecCategory.options[frmSpec.SpecCategory.selectedIndex].value;
document.getElementById("spShow").innerHTML = "_";
</script>
</td>
</tr>

<tr>
<td align=middle width=100 height=30>排序</td>
<td><input class="input" maxlength="6" value="9" style="WIDTH: 60px" name="SpecOrder"/> </td></TR>
<tr>
</TBODY>
</TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="<%if request.QueryString("action")="edit" then%>编辑<%else%>添加<%end if%>" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/>
</DIV></DIV></FORM>

</BODY></HTML>