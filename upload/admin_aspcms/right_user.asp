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
			$("#version").text("网络连接失败");
		}
		if($("#developer").text().replace(" ","")==""){
			$("#developer").text("网络连接失败");
		}
		if($("#developer").text().replace(" ","")=="开源免费版"){
			$("#developer").html("开源免费版 <a href='http://www.aspcms.com/aspcms-3996-1-1.html' style='text-decoration:underline;' target='_blank'>版权声明</a>　<a href='http://www.aspcms.com/aspcms-29177-1-1.html' style='text-decoration:underline;' target='_blank'>购买授权</a>");
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
	<div class="bodytitletxt">安全提示</div>
</div>
<div class="tab_bk1">
	<div class="txtbox1">
		<p class="red">您的后台管理目录为默认目录，请您使用ftp或者其他方式将后台目录重命名</p>
	</div>
</div>
<%end if%>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">网站统计</div>
</div>
<div class="tab_bk1">
	<div class="txtbox1">
		<p>今日 <strong class="red"><%=getTodayVisits%></strong> 人 ，昨日 <strong class="red"><%=getYesterdayVisits%></strong> 人 ，本月 <strong class="red"><%=getMonthVisits%></strong> 人 ，总访问 <strong class="red"><%=getAllVisits%></strong> 人。</p>
	</div>
</div>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">待处理事项</div>
</div>
<div class="tab_bk1">
	<div class="txtbox1">
		<p>
        <a href="_content/_Comments/AspCms_Message.asp">未回答留言</a>（<%=getDataCount("select Count(*) from {prefix}GuestBook where faqStatus=0")%>）<small>|</small> 
        <a href="_content/_Order/AspCms_Order.asp">未处理订单</a>（<%=getDataCount("select Count(*) from {prefix}Order2 where state=0")%>）<small>|</small> 
        <a href="_content/_Comments/AspCms_Comments.asp">未处理评论</a>（<%=getDataCount("select Count(*) from {prefix}Comments where CommentStatus=0")%>）<small>|</small> 
        <a href="_content/_Apply/AspCms_Apply.asp">未处理应聘</a>（<%=getDataCount("select Count(*) from {prefix}Apply where ApplyStatus=0")%>）<small>|</small> 
        </p>
	</div>
</div>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">系统信息</div>
</div>
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1" bgcolor="#cad9ea" >
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft" width="250">当前程序版本：</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%=aspcmsVersion%>     最新版本：<span id="version"></span></td>
	</tr>
    <tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">授权信息</td>
		<td bgcolor="#FFFFFF" class="main_bright">授权版本：V2.2.8　授权用户：<span id="developer"></span></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">服务器名称：</td>
		<td bgcolor="#FFFFFF" class="main_bright">名称 <%=Request.ServerVariables("SERVER_NAME")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>) 端口:<%=Request.ServerVariables("SERVER_PORT")%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">站点物理路径：</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%=Request.ServerVariables("PATH_TRANSLATED")%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">FSO文本读写：</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%If Not isInstallObj(FSO_OBJ_NAME) Then%>
          <font color="#FF0066"><b>×</b>网站不能正常使用</font>
          <%else%>
          <b>√</b>
          <%end if%></td>
	</tr>
    <tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">JMail组件：</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%If Not isInstallObj("JMail.Message") Then%>
        <font color="#FF0066"><b>×</b>邮件提醒功能不能正常使用</font>
        <%else%>
        <b>√</b>
        <%end if%></td>
	</tr>
    <tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft"> ASPJpeg 组件：</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%If Not isInstallObj("Persits.Jpeg") Then%>
        <font color="#FF0066"><b>×</b>水印功能不能正常使用</font>
        <%else%>
        <b>√</b>
        <%end if%></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">脚本解释引擎：</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">服务器IIS版本：</td>
		<td bgcolor="#FFFFFF" class="main_bright"><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	</tr>
</table>


<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">开发团队</div>
</div>
<table width="100%" border="0" align="center" cellpadding="10" cellspacing="1" bgcolor="#cad9ea" >
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft" width="250">版权所有：</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;<strong><a href="http://www.4i.com.cn/" target="_blank">武汉上谷网络科技有限公司(Chancoo Inc.)</a>  </strong></td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">总策划：</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;<a href="http://www.aspcms.com/?1" target="_blank">zhangly</a></td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">开发与支持团队：</td>
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
		<td align="right" bgcolor="#f5fafe" class="main_bleft">界面与用户体验团队：</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://www.aspcms.com/?30062" target="_blank">softocean</a>,  
        <a href="http://www.aspcms.com/?2620" target="_blank">onetwo027</a>, 
        <a href="http://www.aspcms.com/?2699" target="_blank">haoyuanma</a> </td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">感谢贡献者：</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://www.aspcms.com/?2841" target="_blank">phso</a>, 
        <a href="http://www.aspcms.com/?2702" target="_blank">happygiver</a>, 
        <a href="http://www.aspcms.com/?3622" target="_blank">dream200829</a>, 
        <a href="http://www.aspcms.com/?2699" target="_blank">huyaowen</a>, 
        <a href="http://www.aspcms.com/?2633" target="_blank">027</a> </td>
  </tr>
	<tr>
		<td align="right" bgcolor="#f5fafe" class="main_bleft">相关链接：</td>
		<td bgcolor="#FFFFFF" class="main_bright">&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://www.aspcms.com/forum-41-1.html" target="_blank">程序发布</a>,
        <a href="http://www.aspcms.com/forum-38-1.html" target="_blank">商业服务中心</a>,
        <a href="http://www.aspcms.com/forum-46-1.html" target="_blank">问诊求医</a>,
        <a href="http://www.aspcms.com/forum-56-1.html" target="_blank">发展建议/BUG反馈</a><br />&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://www.aspcms.com/forum-47-1.html" target="_blank">模版风格</a>,
        <a href="http://www.aspcms.com/forum-76-1.html" target="_blank">插件/模块</a>,
        <a href="http://www.aspcms.com/forum-51-1.html" target="_blank">站长茶社</a>,
        <a href="http://www.aspcms.com/forum-50-1.html" target="_blank">代码及编程交流(ASP程序员之家)</a></td>
  </tr>
</table>
<div class="bodytitle">
	<div class="bodytitleleft"></div>
	<div class="bodytitletxt">商业信息</div>
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