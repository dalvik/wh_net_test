<!--#include file="AspCms_AdvFun.asp" -->
<%'CheckAdmin("AspCms_AdvAdd.asp")%>
<%CheckAdmin("AspCms_AdvList.asp")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE>��ӹ��</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk"><LINK 
href="../images/style.css" type=text/css rel=stylesheet>
<script type="text/javascript" src="../js/jquery.min.js"></script>
<script type="text/javascript" src="../js/jquery.tinyTips.js"></script>
<script type="text/javascript">
		$(document).ready(function() {
			$('a.tTip').tinyTips('title');
			$('a.imgTip').tinyTips('title');
			$('img.tTip').tinyTips('title');
			$('h1.tagline').tinyTips('tinyTips are totally awesome!');
		});
		</script>
<link rel="stylesheet" type="text/css" media="screen" href="../css/tinyTips.css" />
<script language="javascript" type="text/javascript" src="../js/getdate/WdatePicker.js"></script>
<script>
function ShowBar(type)
{
	var picbar1  = document.getElementById("ipath");
	var picbar2 = document.getElementById("lpath");
	var picbar3 = document.getElementById("cc");
	var picbar4 = document.getElementById("advnr");
	var picbar5 = document.getElementById("ipathr");
	var picbar6 = document.getElementById("lpathr");
	if(type==1 )
	{
		picbar1.style.display="none";
		picbar2.style.display="";
		picbar3.style.display="none";
		picbar4.style.display="";
		picbar5.style.display="none";
		picbar6.style.display="none";
	}
	else if(type==2 || type==4)
	{
		picbar1.style.display="";
		picbar2.style.display="";
		picbar3.style.display="";
		picbar4.style.display="none";
		picbar5.style.display="none";
		picbar6.style.display="none";
	}
	else if(type==3|| type==6)
	{
		picbar1.style.display="none";
		picbar2.style.display="none";
		picbar3.style.display="none";
		picbar4.style.display="";
		picbar5.style.display="none";
		picbar6.style.display="none";
	}
	else if(type==5)
	{
		picbar1.style.display="";
		picbar2.style.display="";
		picbar3.style.display="";
		picbar4.style.display="none";
		picbar5.style.display="";
		picbar6.style.display="";
	}	
	
}
</script> 
</head>
<body>
<FORM id= name= action="" method=post >
<DIV class=formzone>
<DIV class=namezone>��ӹ��</DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>						
    <TD align=middle width=100 height=30>�������</TD>
    <TD><INPUT class="input" style="WIDTH: 200px" maxLength="200" name="advname"/> <FONT color=#ff0000>*</FONT></TD>
  </TR>
 <TR>
    <TD align=middle width=100 height=30>�������</TD>    
    <TD><label>
      <select name="advclass" id="advclass" onChange="ShowBar(this.value)">
      <option value="1">�ı�</option>
      <option value="2">ͼƬ</option>
      <option value="3">����</option>
      </select>
    </label> <img src="../images/help.gif" class="tTip" title="������ͣ�����ģ����Ԥ���ù���λ�ò�������Ӧ��ǩ"></TD>
 </tr>
  <TR id="ipath" style="display:none">
    <TD align=middle width=100 height=30>ͼƬ��ַ</TD>
    <TD><input class="input" style="WIDTH: 300px" name="ImagePath" id="ImagePath" /> <img src="../images/help.gif" class="tTip" title="���ͼƬ"><br>
<iframe src="../editor/upload.asp?sortType=14&stype=file&amp;Tobj=ImagePath" scrolling="no" topmargin="0" height="40" width="100%" marginwidth="0" marginheight="0" frameborder="0" align="left"></iframe> </TD>
  </TR>
  <TR id="lpath">
    <TD align=middle width=100 height=30>���ӵ�ַ</TD>
    <TD><input class="input" style="WIDTH: 300px" name="LinkPath" id="LinkPath" /> <img src="../images/help.gif" class="tTip" title="�ù������ӵ�ַ"></TD>
  </TR>
  <TR id="ipathr" style="display:none">
    <TD align=middle width=100 height=30>�Ҳ�ͼƬ��ַ</TD>
    <TD><input class="input" style="WIDTH: 300px" name="ImagePathr" id="ImagePathr" /><br>
<iframe src="../editor/upload.asp?sortType=14&stype=file&amp;Tobj=ImagePathr" scrolling="no" topmargin="0" height="40" width="100%" marginwidth="0" marginheight="0" frameborder="0" align="left"></iframe></TD>
  </TR>
  <TR id="lpathr"  style="display:none">
    <TD align=middle width=100 height=30>�Ҳ����ӵ�ַ</TD>
    <TD><input class="input" style="WIDTH: 300px" name="LinkPathr" id="LinkPathr" /></TD>
  </TR>
  <TR id="cc" style="display:none">
    <TD align=middle height=30>ͼƬ�ߴ�</TD>
    <TD>��
      <INPUT class="input" style="WIDTH: 30px" maxLength="200" name="imgw"/>
      ��
      <INPUT class="input" style="WIDTH: 30px" maxLength="200" name="imgh"/> <img src="../images/help.gif" class="tTip" title="����ͼƬ�ߴ�"></TD>
  </TR>  
  <TR>
    <TD align=middle width=100 height=30 >Ͷ��ʱ��</TD>
    <TD><input maxlength="20"  name="stime" class="Wdate" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" value="<%=now()%>"/> <img src="../images/help.gif" class="tTip" title="�����Ч�ڣ������ֵĿ�ʼʱ��ͽ���ʱ�䣬������ʱ���治��ʾ"></TD>
  </TR>
  <TR>
    <TD align=middle width=100 height=30>����ʱ��</TD>
    <TD><input maxlength="20"  name="etime" class="Wdate" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" value="2050-1-1 12:00:00"/></TD>
  </TR>
  <TR>
    <TD align=middle height=30>���״̬</TD>
    <TD><input type="radio" name="AdvStatus" value="1" checked />
����
  <input type="radio" name="AdvStatus" value="0" />
���� <img src="../images/help.gif" class="tTip" title="�Ƿ���øù��"></TD>
  </TR>
  <TR id="advnr">
    <TD align=middle height=30>�������</TD>
    <TD>  <TEXTAREA class="textarea" style="WIDTH: 450px" name="advcontent" rows="3"></TEXTAREA>  <img src="../images/help.gif" class="tTip" title="���ֻ�����������"></TD>
  </TR>
    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
  <INPUT class="button" type="submit" value="���" onClick="form.action='?action=add';"/>
<INPUT class="button" type="button" value="����" onClick="history.go(-1)"/>
</DIV>
</DIV></FORM>
</body>
</html>
