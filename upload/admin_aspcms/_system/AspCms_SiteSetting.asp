<!--#include file="AspCms_SettingFun.asp" -->
<%CheckAdmin("AspCms_SiteSetting.asp")%>
<%getComplanySetting%>
<%
'���±�����һ��ʹ��ʱ����������������ᱨ��
on error resume next
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<LINK href="../images/style.css" type=text/css rel=stylesheet>
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
<script type="text/javascript">
        function SetGoogleMapGode(t){
var s = "<";
s += "script src='http://maps.google.com/maps?file=api&amp;v=2&amp;sensor=false&amp;key=";
s += "{aspcms:googlemapkey}' type='text/javascript'>";
s += "</";
s += "script>";
s += "\n";
s += "<";
s += "script src='http://www.google.com/uds/api?file=uds.js&amp;v=1.0' type='text/javascript'>";
s += "</";
s += "script>";
s += "\n";
s += "<";
s += "script src='http://www.google.com/uds/solutions/localsearch/gmlocalsearch.js' type='text/javascript'>";
s += "</";
s += "script>";
s += "\n";
	document.getElementById(t).innerText = s;
	document.getElementById(t).style.display = "block";
		}
        </script>
        <script type="text/javascript">
        $(function(){
			$("input[name='waterType']").change(function(){
					var $type=$("input[name='waterType']:checked").val();
					if($type==0){
						$(".textWater").css("display","block");
						$(".picWater").css("display","none");
					}else{
						$(".textWater").css("display","none");
						$(".picWater").css("display","block");
					}
			});
			$("input[name='waterMark']").change(function(){
				if($(this).val()==0){
					$(".water").css("display","none");
				}else{
					$(".water").css("display","block");
				}
			});
			$("input[name='markPicAlpha']").focus(function(){
				$("#alpha").css("display","inline-block").css("color","#CD0000");
			});
			$("input[name='markPicAlpha']").blur(function(){
				$("#alpha").css("display","none").css("color","#FFFFFF");
			});
		});
        </script>
</HEAD>
<BODY>
<DIV class=formzone>
<FORM name="" action="?action=saves" method="post">
<DIV class=tablezone>

<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
<TBODY>
    <TR>
    <TD width="193" class=innerbiaoti><STRONG>��վ����</STRONG></TD>
        <TD class=innerbiaoti width=568 height=28></TD>
    </TR>
    <TR class=list>
        <TD>��վ����ģʽ : </TD>
        <TD width=568>
        <input type="radio" name="runMode" value="0" <% if runMode=0 then echo"checked" end if %>>��̬
        <input type="radio" name="runMode" value="1" <% if runMode=1 then echo"checked" end if %>>��̬   <img src="../images/help.gif" class="tTip" title="���ѡ��̬ģʽ�����ڸ�����Ϣ���ھ�̬�������������༭����Ϣ"></TD>
    </TR>
    <TR class=list>
        <TD>��վ״̬ : </TD>
        <TD width=568>
        <input type="radio" name="siteMode" value="1" <% if siteMode=1 then echo"checked" end if %>>����
        <input type="radio" name="siteMode" value="0" <% if siteMode=0 then echo"checked" end if %>>�ر�        </TD>
    </TR>
    <TR class=list>
        <TD>��վ�ر�˵�� : </TD>
        <TD><INPUT class="input" size="30" style="width:80%" value="<%=siteHelp%>" name="siteHelp"/> </TD>
    </TR>
    <TR class=list>
      <TD>��վ��̨��֤�뿪��</TD>
      <TD><input type="radio" value="1" name="admincode" <% if admincode=1 then echo "checked" end if%>>����
        <input type="radio" value="0" name="admincode" <% if admincode=0 then echo "checked" end if%>>�ر�
          </TD>
    </TR>
    <TR class=list>
        <TD>���۹��ܿ��� : </TD>
        <TD>
        	<input type="radio" value="1" name="SwitchComments" <% if SwitchComments=1 then echo "checked" end if%>>����
            <input type="radio" value="0" name="SwitchComments" <% if SwitchComments=0 then echo "checked" end if%>>�ر� <img src="../images/help.gif" class="tTip" title="���ۿ����Լ�������˿�����Ҫ��ҳ���е�����Ӧ�����۱�ǩ��������">        </TD>
    </TR>
    <TR class=list>
        <TD>������˿��� : </TD>
        <TD><input type="radio" value="1" name="SwitchCommentsStatus" <% if SwitchCommentsStatus=1 then echo "checked" end if%>>����
            <input type="radio" value="0" name="SwitchCommentsStatus" <% if SwitchCommentsStatus=0 then echo "checked" end if%>>�ر�</TD>
    </TR>
    <TR class=list>
        <TD>���Թ��ܿ��� : </TD>
        <TD>
            <input type="radio" value="1" name="switchFaq" <% if switchFaq=1 then echo "checked" end if%>>����
            <input type="radio" value="0" name="switchFaq" <% if switchFaq=0 then echo "checked" end if%>>�ر�         </TD>
    </TR> <TR class=list>
        <TD>������˿��� : </TD>
        <TD>
            <input type="radio" value="1" name="SwitchFaqStatus" <% if SwitchFaqStatus=1 then echo "checked" end if%>>����
            <input type="radio" value="0" name="SwitchFaqStatus" <% if SwitchFaqStatus=0 then echo "checked" end if%>>�ر�  <img src="../images/help.gif" class="tTip" title="���Կ����Լ�������˿�����Ҫ��ҳ���е�����Ӧ�����Ա�ǩ��������">         </TD>
    </TR>
    <TR class=list>
        <TD>�������ÿ��� : </TD>
        <TD><input type="radio" value="1" name="dirtyStrToggle" <% if dirtyStrToggle=1 then echo "checked" end if%>>����
            <input type="radio" value="0" name="dirtyStrToggle" <% if dirtyStrToggle=0 then echo "checked" end if%>>�ر� <img src="../images/help.gif" class="tTip" title="��վ�й��˵��Ƿ��ַ������ڹ���������������Ҫ���˵��ַ�"></TD>
    </TR>
    <TR class=list>
        <TD>�������� : </TD>
        <TD><textarea class="textarea" name="dirtyStr" cols="60" style="width:98%" rows="8"><%=decode(dirtyStr)%></textarea></TD>
    </TR>

    <TR>
    	<TD class=innerbiaoti><STRONG>ˮӡ����</STRONG></TD>
        <TD class=innerbiaoti></TD>
    </TR> 
    
    <TR class=list>
        <TD>�Ƿ���ˮӡ���� : </TD>
        <TD>
        <input type="radio" name="waterMark" value="1" <% if waterMark=1 then echo"checked" end if %>>����
        <input type="radio" name="waterMark" value="0" <% if waterMark=0 then echo"checked" end if %>>�ر� <img src="../images/help.gif" class="tTip" title="ˮӡ������Ҫ��������֧�֣�ֻ�ܶ�ͼƬ���м�ˮӡ"></TD>
    </TR>
   <TR class="list water">
        <TD>ˮӡ���� : </TD>
        <TD>
        <input type="radio" name="waterType" value="0" <% if waterType=0 then echo"checked" end if %>>����
        <input type="radio" name="waterType" value="1" <% if waterType=1 then echo"checked" end if %>>ͼƬ
        </TD>
    </TR>
    <TR class="list textWater water"  style="<% if waterType=0 then %>display:block<%else%>display:none<%end if %>">
        <TD>ˮӡ���� : </TD>
        <TD><input class="input" size="30" value="<%=waterMarkFont%>" name="waterMarkFont"/></TD>
    </TR>
    <tr class="list picWater water" style="<% if waterType=1 then %>display:block<%else%>display:none<%end if %>">
    	<td>ͼƬ·����</td>
        <td><input type="text" value="<%=waterMarkPic%>" name="waterMarkPic" id="indeximage" style="width:300px;"><br /><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../editor/upload.asp?sortType=99&stype=file&Tobj=indeximage" allowTransparency="true"></iframe></td>
    </tr>
    <!--<tr class="list picWater water" type="<%' if waterType=1 then %>display:block<%'else%>display:none<%'end if %>">
    	<td>ͼƬ���ã�</td>
        <td>��ȣ�<input  type="text" name="markPicWidth" value="<%'=markPicWidth%>" size="4" style="vertical-align:middle">�� �߶ȣ�<input type="text" me="markPicHeight" value="<%'=markPicHeight%>" size="4"  style="vertical-align:middle">�� ͸���ȣ�<input type="text" value="<%'=markPicAlpha%>" size="4" name="markPicAlpha"  style="vertical-align:middle"><span id="alpha" style="display:none">(������С�������Ϊ1����СΪ0��Ĭ��Ϊ0.5)</span></td>
    </tr>-->
    <TR class="list water">
        <TD>ˮӡλ�� : </TD>
        <TD>
          <table width="300" border="0" cellspacing="1" cellpadding="0">
            <tr>
                <td width="33%"><input type="radio" name="waterMarkLocation"  value="1" <% if waterMarkLocation=1 then echo"checked" end if %>>��������</td>
                <td width="33%"><input type="radio" name="waterMarkLocation"  value="2" <% if waterMarkLocation=2 then echo"checked" end if %>>��������</td>
                <td><input type="radio" name="waterMarkLocation"  value="3" <% if waterMarkLocation=3 then echo"checked" end if %>>��������</td>
            </tr>
            <tr>
                <td><input type="radio" name="waterMarkLocation"  value="4" <% if waterMarkLocation=4 then echo"checked" end if %>>��߾���</td>
                <td><input type="radio" name="waterMarkLocation"  value="5" <% if waterMarkLocation=5 then echo"checked" end if %>>ͼƬ����</td>
                <td><input type="radio" name="waterMarkLocation"  value="6" <% if waterMarkLocation=6 then echo"checked" end if %>>�ұ߾���</td>
            </tr>
            <tr>
                <td><input type="radio" name="waterMarkLocation"  value="7" <% if waterMarkLocation=7 then echo"checked" end if %>>�ײ�����</td>
                <td><input type="radio" name="waterMarkLocation"  value="8" <% if waterMarkLocation=8 then echo"checked" end if %>>�ײ�����</td>
                <td><input name="waterMarkLocation" type="radio"  value="9" <% if waterMarkLocation=9 then echo"checked" end if %>>�ײ�����</td>
            </tr>
         </table>
         </TD>
    </TR>
    
    <TR>
    	<TD class=innerbiaoti><STRONG>�ʼ�����</STRONG></TD>
        <TD class=innerbiaoti></TD>
    </TR>   
    <TR class=list>
        <TD>�ʼ���ַ : </TD>
        <TD><INPUT class="input" size="30" value="<%=smtp_usermail%>" name="smtp_usermail"/> </TD>
    </TR>
    <TR class=list>
        <TD>�ʼ��˺� : </TD>
        <TD><INPUT class="input" size="30" value="<%=smtp_user%>" name="smtp_user"/> </TD>
    </TR>
    <TR class=list>
        <TD>�ʼ����� : </TD>
        <TD><INPUT class="input" size="30" value="<%=smtp_password%>" name="smtp_password"/> </TD>
    </TR>
    <TR class=list>
        <TD>�ʼ������� : </TD>
        <TD><INPUT class="input" size="30" value="<%=smtp_server%>" name="smtp_server"/> </TD>
    </TR>
    <TR>
    	<TD class=innerbiaoti><STRONG>���ѹ���</STRONG></TD>
        <TD class=innerbiaoti></TD>
    </TR>   
    <TR class=list>
        <TD>�ʼ���ַ : </TD>
        <TD><INPUT class="input" size="30" value="<%=MessageAlertsEmail%>" name="MessageAlertsEmail"/> </TD>
    </TR>
    <TR class=list>
        <TD>�������� : </TD>
        <TD>
        <input type="radio" name="messageReminded" value="1" <% if messageReminded=1 then echo"checked" end if %>>����
        <input type="radio" name="messageReminded" value="0" <% if messageReminded=0 then echo"checked" end if %>>�ر�  </TD>
    </TR>
    <TR class=list>
        <TD>�������� : </TD>
        <TD>
        <input type="radio" name="orderReminded" value="1" <% if orderReminded=1 then echo"checked" end if %>>����
        <input type="radio" name="orderReminded" value="0" <% if orderReminded=0 then echo"checked" end if %>>�ر�  </TD>
    </TR>
    <TR class=list>
        <TD>ӦƸ���� : </TD>
        <TD>
        <input type="radio" name="applyReminded" value="1" <% if applyReminded=1 then echo"checked" end if %>>����
        <input type="radio" name="applyReminded" value="0" <% if applyReminded=0 then echo"checked" end if %>>�ر�  </TD>
    </TR>
    <TR class=list>
        <TD>�������� : </TD>
        <TD>
        <input type="radio" name="commentReminded" value="1" <% if commentReminded=1 then echo"checked" end if %>>����
        <input type="radio" name="commentReminded" value="0" <% if commentReminded=0 then echo"checked" end if %>>�ر�  </TD>
    </TR>
</TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT type="hidden" value="<%=LanguageID%>" name="LanguageID" />
<INPUT class="button" type="submit" value="����" />
<INPUT class="button" type="button" value="����" onClick="history.go(-1)"/> 
 </DIV></FORM>
</DIV>
</BODY>
</HTML>
