<!--#include file="../../../inc/AspCms_MainClass.asp" -->
<%
CheckAdmin("AspCms_Slide.asp")
dim action : action=getForm("action", "get")
Select Case action
	Case "save"	 : Call saveService()
End Select 
	
Sub saveService
	dim templateobj,configStr,configPath
	
	
	if not isnum(getForm("slideNum","post")) then alertMsgAndGo "�õ�Ƭ��������ȷ","-1"
	configPath="../../../config/AspCms_Config.asp"
	set templateobj=new TemplateClass
	configStr=loadFile(configPath)
	configStr=templateobj.regExpReplace(configStr,"Const slideTextStatusB=(\d?)","Const slideTextStatusB="&getForm("slideTextStatus","post")&"")
	configStr=templateobj.regExpReplace(configStr,"Const slideNumB=([\d]*)","Const slideNumB="&getForm("slideNum","post")&"")
	
	configStr=templateobj.regExpReplace(configStr,"Const slideWidthB=""(\S*?)""","Const slideWidthB="""&getForm("slideWidth","post")&"""")
	
	configStr=templateobj.regExpReplace(configStr,"Const slideHeightB=""(\S*?)""","Const slideHeightB="""&getForm("slideHeight","post")&"""")
	configStr=templateobj.regExpReplace(configStr,"Const slideTextsB=""([\s\S]*?)""","Const slideTextsB="""&setText(getForm("slideNum","post"),getForm("slideTexts","post"))&"""")
	configStr=templateobj.regExpReplace(configStr,"Const slideLinksB=""([\s\S]*?)""","Const slideLinksB="""&setText(getForm("slideNum","post"),getForm("slideLinks","post"))&"""")
	configStr=templateobj.regExpReplace(configStr,"Const slidestyleB=(\d?)","Const slidestyleB="&getForm("slidestyle","post")&"")
	configStr=templateobj.regExpReplace(configStr,"Const slideImgsB=""([\s\S]*?)""","Const slideImgsB="""&setText(getForm("slideNum","post"),getForm("slideImgs","post"))&"""")
	'die configStr
	'die "Const slideImgs="""&setText(getForm("slideNum","post"),getForm("slideImgs","post"))&""""
	
	createTextFile configStr,configPath,""
	set templateobj=nothing
	alertMsgAndGo "�޸ĳɹ�",getPageName()
	
End Sub


Function setText(num,text)
	dim i 
	text=split(text,",")
	for i=0 to num-1
		if ubound(text)+1>i then
			setText=setText&text(i)&","
		else			
			setText=setText&","
		end if	
	next
End Function


%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>
</HEAD>
<BODY>
<DIV class=searchzone>
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD height=30>�õ�Ƭ����B</TD>
    <TD align=right colSpan=2><span class="piliang">
     <INPUT onClick="location.href='AspCms_Slide.asp'" type="button" value="�õ�ƬA" class="button" /> <INPUT onClick="location.href='AspCms_SlideB.asp'" type="button" value="�õ�ƬB" class="button" /> <INPUT onClick="location.href='AspCms_SlideC.asp'" type="button" value="�õ�ƬC" class="button" /> <INPUT onClick="location.href='AspCms_SlideD.asp'" type="button" value="�õ�ƬD" class="button" />
    </span></TD>
  </TR></TBODY></TABLE></DIV>
<FORM name="form" action="?action=save" method="post" >
<DIV class=formzone>
  <DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>						
    <TD align="middle" width=100 height=30>ͼƬ����</TD>
    <TD align="left">
    <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#a9c5d0" id="">
              <tr bgcolor="#DDEFF9" align="center">
                <td width="31%" height="28">ͼƬ</td>
                <td width="69%">�ϴ�</td>
              </tr>
			<%
            Dim slideImg,slideLink,slideText,i,imgstr
            slideImg = split(slideImgsB,",")
            slideLink = split(slideLinksB,",")
            slideText = split(slideTextsB,",")
            
            for i=0 to slideNumB-1			
				
					imgstr="<img id=img"&i&" src="""&trim(slideImg(i))&""" width=""160px"" height=""120px"" />"
							
            %>
              <tr align="center" bgcolor="#FFFFFF">
            	<td class="pic"><%=imgstr%></td>
                <td align="left"><table width="95%" border="0" cellspacing="2" cellpadding="0">                
                  <tr>
                    <td width="18%" align="right">ͼƬ��ַ��</td>
                    <td width="82%"><input class="input" style="width:300px" type="text" name="slideImgs" id="ImagePath<%=i%>" value="<%=trim(slideImg(i))%>"/></td>
                  </tr>
                  <tr>
                    <td align="right" valign="top">�ϴ���</td>
                    <td><iframe src="../../editor/upload.asp?sortType=12&stype=file&Tobj=ImagePath<%=i%>&toimg=img<%=i%>" scrolling="No" topmargin="0" width="100%" height="40" marginwidth="0" marginheight="0" frameborder="0" align="left"></iframe></td>
                  </tr>
                  <tr>
                    <td align="right">���ӵ�ַ��</td>
                    <td><input type="text" class="input" style="width:60%" name="slideLinks"  value="<%=trim(slideLink(i))%>"/>����������"#"</td>
                  </tr>
                  <tr>
                    <td align="right">˵�����֣� </td>
                    <td><input type="text" class="input" style="width:80%" name="slideTexts"  value="<%=trim(slideText(i))%>"/></td>
                  </tr>
                </table></td>
              </tr>
              <%
			  next
			  %>                        
            </table>    </TD>
  </TR>
  <TR>						
    <TD align=middle width=100 height=30>�õ�Ƭ����</TD>
    <TD><INPUT class="input" style="WIDTH: 40px" maxLength="2" name="slideNum" value="<%=slideNumB%>"/></TD>
  </TR>
  <TR>
    <TD align=middle width=100 height=30>�õ�Ƭ��</TD>    
    <TD><input type="text" class="input" name="slideHeight" value="<%=slideHeightB%>" style="width:40px" />px</TD>
</tr>
  <TR>
    <TD align=middle width=100 height=30>�õ�Ƭ��</TD>    
    <TD>
		 <input type="text" class="input" name="slideWidth" value="<%=slideWidthB%>" style="width:40px" />px</TD>
</tr>

  <TR>
    <TD align=middle width=100 height=30>����˵��</TD>
    <TD>
	 <input type="radio"  value="1" name="slideTextStatus" <% if slideTextStatusB=1 then echo "checked" end if%>/>��ʾ
            <input type="radio" value="0" name="slideTextStatus" <% if slideTextStatusB=0 then echo "checked" end if%>/>����</TD></TR>
  <TR>
  <TR>
    <TD align=middle height=30>�õ���ʽ</TD>
    <TD><input type="radio"  value="0" name="slidestyle" <% if slidestyleB=0 then echo "checked" end if%>/>Ĭ��
            <input type="radio" value="1" name="slidestyle" <% if slidestyleB=1 then echo "checked" end if%>/><img src="../../images/slide1.jpg" width="150"></TD>
  </TR>
  <TR>
    <TD align=middle width=100 height=30>�õ�Ƭ���ñ�ǩ</TD>
    <TD>{aspcms:slideb}</TD></TR>
  <TR>
    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="����" />
<INPUT class="button" type="button" value="����" onClick="history.go(-1)"/> 
<INPUT onClick="location.href='<%=getPageName()%>'" type="button" value="ˢ��" class="button" /> 
</DIV></DIV></FORM>

</BODY></HTML>