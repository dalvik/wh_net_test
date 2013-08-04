<!--#include file="AspCms_SpecFun.asp" -->
<%CheckAdmin("AspCms_Spec.asp")%>
<%EditSpec%>
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
<FORM action="?action=editsave" method="post" name="frmSpec">
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

<TR style="display:none">
<td align=middle width=100 height=30>应用与分类</td>    
<td>
<input name="SpecCategory" type="hidden" value="<%=SpecCategory%>">
</td>
</tr>

<tr  style="display:none">
<td align=middle width=100 height=30>排序</td>
<td><input class="input" maxlength="6" value="9" style="WIDTH: 60px" name="SpecOrder"/> </td></TR>
<tr>
</TBODY>
</TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="修改" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/>
</DIV></DIV></FORM>

</BODY></HTML>