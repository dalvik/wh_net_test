<!--#include file="AspCms_TabFun.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>

<script type="text/javascript">
function changeTabControlType(select){
	var trTabOptions = document.getElementById("trTabOptions");
	if(select.options[select.selectedIndex].value>5)
		trTabOptions.style.display = "block";
	else
		trTabOptions.style.display = "none";
}
function changeTabCategoryType(select){
	var preShow = document.getElementById("preShow");
	var spShow  = document.getElementById("spShow");
	preShow.innerHTML = select.options[select.selectedIndex].value;
	if(spShow.innerHTML == "")spShow.innerHTML = "_";
	
}
</script>
</HEAD>
<BODY>
<FORM action="?action=<%if request.QueryString("action")="edit" then%>editsave<%else%>add<%end if%>" method="post" name="frmTab">
<input type="hidden" name="TabDiversification" value="<%=TabDiversification%>" />
<input type="hidden" name="TabID" value="<%=TabID%>" />
<DIV class=formzone>
<DIV class=namezone><%if request.QueryString("action")="edit" then%>�༭����<%else%>��Ӳ���<%end if%></DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
<TBODY>
<TR>						
<TD align=middle width=100 height=30>��������</TD>
<TD><INPUT class="input" style="WIDTH: 200px" maxLength="200" name="TabName" value="<%=TabName%>"/> <FONT color=#ff0000>*</FONT> </TD>
</TR>
<TR>
<TD align=middle width=100 height=30>�ֶ�����</TD>    
<TD><INPUT class="input" style="WIDTH: 200px" onKeyUp="document.getElementById('endShow').innerHTML=value" maxLength="200" name="TabField"  value="<%=TabField%>"/> <FONT color=#ff0000>*Ԥ�� [<span id="preShow"><%=TabCategory%></span><span id="spShow"><%if TabCategory <> "" then%>_<%end if%></span><span id="endShow"><%=TabField%></span>]</FONT></TD>
</tr>
<TR style="display:none">
<TD align=middle width=100 height=30>������д</TD>    
<TD><INPUT type="checkbox" class="input" name="TabNotNull" <%if TabNotNull then%>checked<%end if%> /></TD>
</tr>
<TR style="display:none">
<TD align=middle width=100 height=30 >�ؼ�����</TD>    
<TD><select name="TabControlType" onChange="changeTabControlType(this)">
<option value="0" selected>�ı�</option>
<option value="1">����</option>
<option value="2">�ı�����</option>
<option value="4">����</option>
<option value="5">��ɫ</option>
<option value="6">��ѡ</option>
<option value="7">��ѡ</option>
</select>
<script>
//frmTab.TabControlType.selectedIndex = <%=TabControlType%>;
for(var i=0;i<frmTab.TabControlType.options.length;i++){
	if(frmTab.TabControlType.options[i].value == "<%=TabControlType%>")
		frmTab.TabControlType.selectedIndex = i;
}
</script>
</TD>
</tr>
<TR style="display:none" id="trTabOptions">
<TD align=middle width=100 height=30>��ѡ����</TD>    
<TD><textarea name="TabOptions" cols="50" rows="5" class="input" style="width:500px;height:80px;"><%=TabOptions%></textarea>�������ûس�����</TD>
</tr>
<TR>

</tr>

<tr>
<td align=middle width=100 height=30>����</td>
<td><input class="input" maxlength="6" value="9" style="WIDTH: 60px" name="TabOrder"/> </td></TR>
<tr>
</TBODY>
</TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="<%if request.QueryString("action")="edit" then%>�༭<%else%>���<%end if%>" />
<INPUT class="button" type="button" value="����" onClick="history.go(-1)"/>
</DIV></DIV></FORM>

</BODY></HTML>