<!--#include file="inc/AspCms_SettingClass.asp" -->
<%CheckLogin()%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Home</title>
<link href="css/css_body.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/js/jquery.js"></script>
<script type="text/javascript">
	$(function(){
		if($("#version").text().replace(" ","")==""){
			$("#version").text("��������ʧ��");
		}
		if($("#developer").text().replace(" ","")==""){
			$("#developer").text("��������ʧ��");
		}
		if($("#developer").text().replace(" ","")=="��Դ��Ѱ�"){
			$("#developer").html("��Դ��Ѱ� <a href='http://www.aspcms.com/aspcms-3996-1-1.html' style='text-decoration:underline;' target='_blank'>��Ȩ����</a>��<a href='http://www.aspcms.com/aspcms-29177-1-1.html' style='text-decoration:underline;' target='_blank'>������Ȩ</a>");
		}
	})
</script>
</script>
</head>

<body>
<%
if "admin"=lcase(getAdminDir) then 
%>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">��ȫ��ʾ</div>
</div>
<div class="tab_bk1">
	<div class="txtbox1">
		<p class="red">���ĺ�̨����Ŀ¼ΪĬ��Ŀ¼������ʹ��ftp����������ʽ����̨Ŀ¼������</p>
	</div>
</div>
<%end if%>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">��վͳ��</div>
</div>
<div class="tab_bk1">
	<div class="txtbox1">
		<p>���� <strong class="red"><%=getTodayVisits%></strong> �� ������ <strong class="red"><%=getYesterdayVisits%></strong> �� ������ <strong class="red"><%=getMonthVisits%></strong> �� ���ܷ��� <strong class="red"><%=getAllVisits%></strong> �ˡ�</p>
	</div>
</div>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">����������</div>
</div>
<div class="tab_bk1">
	<div class="txtbox1">
		<p>
        <a href="_content/_Comments/AspCms_Message.asp">δ�ش�����</a>��<%=getDataCount("select Count(*) from {prefix}GuestBook where faqStatus=0")%>��<small>|</small> 
        <a href="_content/_Order/AspCms_Order.asp">δ������</a>��<%=getDataCount("select Count(*) from {prefix}Order2 where state=0")%>��<small>|</small> 
        <a href="_content/_Comments/AspCms_Comments.asp">δ��������</a>��<%=getDataCount("select Count(*) from {prefix}Comments where CommentStatus=0")%>��<small>|</small> 
        <a href="_content/_Apply/AspCms_Apply.asp">δ����ӦƸ</a>��<%=getDataCount("select Count(*) from {prefix}Apply where ApplyStatus=0")%>��<small>|</small> 
        </p>
	</div>
</div>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">ϵͳ��Ϣ</div>
</div>
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1" bgcolor="#cad9ea" >
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft" width="250">��ǰ����汾��</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%=aspcmsVersion%>     ���°汾��<span id="version"></span></td>
	</tr>
    <tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">��Ȩ��Ϣ</td>
		<td bgcolor="#FFFFFF" class="main_bright">��Ȩ�汾��V2.2.8����Ȩ�û���<span id="developer"></span></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">���������ƣ�</td>
		<td bgcolor="#FFFFFF" class="main_bright">���� <%=Request.ServerVariables("SERVER_NAME")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>) �˿�:<%=Request.ServerVariables("SERVER_PORT")%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">վ������·����</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%=Request.ServerVariables("PATH_TRANSLATED")%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">FSO�ı���д��</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%If Not isInstallObj(FSO_OBJ_NAME) Then%>
          <font color="#FF0066"><b>��</b>��վ��������ʹ��</font>
          <%else%>
          <b>��</b>
          <%end if%></td>
	</tr>
    <tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">JMail�����</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%If Not isInstallObj("JMail.Message") Then%>
        <font color="#FF0066"><b>��</b>�ʼ����ѹ��ܲ�������ʹ��</font>
        <%else%>
        <b>��</b>
        <%end if%></td>
	</tr>
    <tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft"> ASPJpeg �����</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%If Not isInstallObj("Persits.Jpeg") Then%>
        <font color="#FF0066"><b>��</b>ˮӡ���ܲ�������ʹ��</font>
        <%else%>
        <b>��</b>
        <%end if%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">�ű��������棺</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">������IIS�汾��</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	</tr>
</table>


<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">�����Ŷ�</div>
</div>
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1" bgcolor="#cad9ea" >
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft" width="250">��Ȩ���У�</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;<strong><a href="http://www.4i.com.cn/" target="_blank">�人�Ϲ�����Ƽ����޹�˾(Chancoo Inc.)</a>  </strong></td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">�ܲ߻���</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://www.aspcms.com/?1" target="_blank">zhangly</a></td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">������֧���Ŷӣ�</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;
        
         <a href="http://www.aspcms.com/?45479" target="_blank">wuya1235</a>,
         <a href="http://www.aspcms.com/?43890">sgkj</a>,  
        
        <a href="http://www.aspcms.com/?2995" target="_blank">enjoyaspcms</a>, 
        <a href="http://www.aspcms.com/?2" target="_blank">zero</a>,  
        <a href="http://www.aspcms.com/?7397" target="_blank">aspcmser</a>, 
        <a href="http://www.aspcms.com/?8901" target="_blank">xiongxt</a>, 
        <a href="http://www.aspcms.com/?3" target="_blank">admin</a></td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">�������û������Ŷӣ�</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://www.aspcms.com/?30062" target="_blank">softocean</a>,  
        <a href="http://www.aspcms.com/?2620" target="_blank">onetwo027</a>, 
        <a href="http://www.aspcms.com/?2699" target="_blank">haoyuanma</a> </td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">��л�����ߣ�</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://www.aspcms.com/?2841" target="_blank">phso</a>, 
        <a href="http://www.aspcms.com/?2702" target="_blank">happygiver</a>, 
        <a href="http://www.aspcms.com/?3622" target="_blank">dream200829</a>, 
        <a href="http://www.aspcms.com/?2699" target="_blank">huyaowen</a>, 
        <a href="http://www.aspcms.com/?2633" target="_blank">027</a> </td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">������ӣ�</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://www.aspcms.com/forum-41-1.html" target="_blank">���򷢲�</a>,
        <a href="http://www.aspcms.com/forum-38-1.html" target="_blank">��ҵ��������</a>,
        <a href="http://www.aspcms.com/forum-46-1.html" target="_blank">������ҽ</a>,
        <a href="http://www.aspcms.com/forum-56-1.html" target="_blank">��չ����/BUG����</a><br />&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://www.aspcms.com/forum-47-1.html" target="_blank">ģ����</a>,
        <a href="http://www.aspcms.com/forum-76-1.html" target="_blank">���/ģ��</a>,
        <a href="http://www.aspcms.com/forum-51-1.html" target="_blank">վ������</a>,
        <a href="http://www.aspcms.com/forum-50-1.html" target="_blank">���뼰��̽���(ASP����Ա֮��)</a></td>
  </tr>
</table>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">��ҵ��Ϣ</div>
</div>
<div class="tab_bk1">
	<div id="linklist"></div>
    
    <script type="text/javascript" src="http://iiii.aspcms.com/js/adlinkjs.js" /></script>
</div>
<%
dim host
host = request.ServerVariables("HTTP_HOST")
%>
<script type="text/javascript" src='http://iiii.aspcms.com/deveplor.asp?url=<%=host%>'></script>
<script type="text/javascript" src='http://iiii.aspcms.com/version.asp?ver=<%=currentversion%>&url=<%=host%>'></script>
</body>
</html>