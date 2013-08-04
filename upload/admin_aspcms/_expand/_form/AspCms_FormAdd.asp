<!--#include file="AspCms_FormFun.asp" -->
<%CheckAdmin("AspCms_Custom.asp")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>

<script type="text/javascript">
function changeFormControlType(select){
	var trFormOptions = document.getElementById("trFormOptions");
	if(select.options[select.selectedIndex].value>5)
		trFormOptions.style.display = "block";
	else
		trFormOptions.style.display = "none";
}
function changeFormCategoryType(select){
	var preShow = document.getElementById("preShow");
	var spShow  = document.getElementById("spShow");
	preShow.innerHTML = select.options[select.selectedIndex].value;
	if(spShow.innerHTML == "")spShow.innerHTML = "_";
	
}
</script>
</HEAD>
<BODY>
<FORM action="?action=<%if request.QueryString("action")="edit" then%>editsave<%else%>add<%end if%>" method="post" name="frmForm">
<input type="hidden" name="FormDiversification" value="<%=FormDiversification%>" />
<input type="hidden" name="FormID" value="<%=FormID%>" />
<DIV class=formzone>
<DIV class=namezone><%if request.QueryString("action")="edit" then%>编辑参数<%else%>添加参数<%end if%></DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
<TBODY>
<TR>						
<TD align=middle width=100 height=30>参数名称</TD>
<TD><INPUT class="input" style="WIDTH: 200px" maxLength="200" name="FormName" value="<%=FormName%>"/> <FONT color=#ff0000>*</FONT> </TD>
</TR>
<TR>
<TD align=middle width=100 height=30>字段名称</TD>    
<TD><INPUT class="input" style="WIDTH: 200px" onKeyUp="document.getElementById('endShow').innerHTML=value" maxLength="200" name="FormField"  value="<%=FormField%>"/> <FONT color=#ff0000>*预览 [<span id="preShow"><%=FormCategory%></span><span id="spShow"><%if FormCategory <> "" then%>_<%end if%></span><span id="endShow"><%=FormField%></span>]</FONT></TD>
</tr>
<TR style="display:none">
<TD align=middle width=100 height=30 >必须填写</TD>    
<TD><INPUT type="checkbox" class="input" name="FormNotNull" <%if FormNotNull then%>checked<%end if%> /></TD>
</tr>
<TR style="display:none">
<TD align=middle width=100 height=30>控件类型</TD>    
<TD><select name="FormControlType" onChange="changeFormControlType(this)">
<option value="0" selected>文本</option>

</select>
<script>
//frmTab.TabControlType.selectedIndex = <%=FormControlType%>;
for(var i=0;i<frmForm.FormControlType.options.length;i++){
	if(frmForm.FormControlType.options[i].value == "<%=FormControlType%>")
		frmForm.FormControlType.selectedIndex = i;
}
</script>
</TD>
</tr>
<TR style="display:none" id="trFormOptions">
<TD align=middle width=100 height=30>备选内容</TD>    
<TD><textarea name="FormOptions" cols="50" rows="5" class="input" style="width:500px;height:80px;"><%=FormOptions%></textarea>多条请用回车区别</TD>
</tr>
<TR>

</tr>

<tr>
<td align=middle width=100 height=30>排序</td>
<td><input class="input" maxlength="6" value="9" style="WIDTH: 60px" name="FormOrder"/> </td></TR>
<tr>
</TBODY>
</TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="<%if request.QueryString("action")="edit" then%>编辑<%else%>添加<%end if%>" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/>
</DIV></DIV></FORM>

</BODY></HTML>