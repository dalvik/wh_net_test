<!--#include file="inc/AspCms_SettingClass.asp" -->
<%CheckLogin()%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Top</title>
<link href="css/css_top_user.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="js/menu.js"></script>
</head>
<body>
<div class="topnav">
	<div class="sitenav">
		<div class="welcome">
			��ã�<span style="color:#F00"><%=session("GroupName")%></span><span class="username"><%=session("adminName")%></span>
            <a href="editPass.asp" target="main">�޸�����</a>
			<a href="login.asp?action=logout" target="_parent">ע����¼</a>
		</div>
		<div class="resize">
		<a href="javascript:ChangeMenu(-1)"><img src='images/frame-l.gif' border='0' alt="��С����"></a>
    <a href="javascript:ChangeMenu(0)"><img src='images/frame_on.gif' border='0' alt="����/��ʾ����"></a>
    <a href="javascript:ChangeMenu(1)" title="��������"><img src='images/frame-r.gif' border='0' alt="��������"></a>
    </div>
		<div class="sitelink">
			<a href="right.asp" target="main">��̨��ҳ</a>
			<a href="http://b.aspcms.com/list/?24_1.html" target="_blank">��������</a>
			<a href="../" target="_blank">��վ��ҳ</a>
	
		</div>
	</div>
	<div class="leftnav">
		<ul>
			<li class="navleft"></li>
<li id='d1' class=""thisclass""><a href="_content/_Comments/AspCms_Message.asp" target="main">���Թ���</a></li>
<li id='d2' class=""thisclass""><a href="_content/_Order/AspCms_Order.asp" target="main">��������</a></li>
<li id='d3' class=""thisclass""><a href="_content/_Apply/AspCms_Apply.asp" target="main">ӦƸ����</a></li>
<li id='d4' class=""thisclass""><a href="_seo/AspCms_MakeHtml.asp?actType=html" target="main">��̬����</a></li>
			<li class="navright"></li>
		</ul>
	</div>
</div>
</body>
</html>
