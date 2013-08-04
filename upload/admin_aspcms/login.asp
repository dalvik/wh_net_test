<!--#include file="inc/AspCms_SettingClass.asp" -->
<%
dim action : action=getForm("action","get")
if action = "login" then
	dim UserName,Password,sql,code,Rs,atype
	UserName = filterPara(getForm("username","post"))
	atype = getForm("atype","post")
	Password = md5(getForm("Password","post"),16)
	code     = getForm("code","post")
	if admincode=1 then
		if code <> Session("Code") then	alertMsgAndGo "您填写的验证码错误!","-1"
	end if
	if isOutSubmit then  alertMsgAndGo"非法外部提交被禁止","-1"

'	User UserGroup
'	LoginName
'Password
'GroupID
'IsAdmin
	sql = "select count(*) from {prefix}User where LoginName = '"& UserName &"' and Password='"&Password&"'"
	Dim rsObj : Set rsObj=Conn.Exec(sql,"r1")
	if rsObj(0)=1 then
	
		
		Set rsObj=Conn.Exec("select IsAdmin, GroupStatus,GroupName, UserStatus, UserID, GroupMenu,GroupSort, LanguageID,adminrand from {prefix}User as a, {prefix}UserGroup as b where LoginName='"&UserName&"' and a.GroupID=b.GroupID","r1")	
		if not rsObj.Eof Then
			if rsObj("IsAdmin")<>1 then alertMsgAndGo"对不起，你不是管理员！","-1"
			if rsObj("GroupStatus")<>1 then alertMsgAndGo"对不起，您所在用户组已被禁用！","-1"
			if rsObj("UserStatus")<>1 then alertMsgAndGo"对不起，您的账号已被禁用！","-1"
			Session("adminName")=UserName
			Session("GroupName")=rsObj("GroupName")
			'wCookie"adminpwd",getForm("Password","post")
			Session("adminpwd") = getForm("Password","post")
			Session("adminId")=rsObj("UserID")
			Session("groupMenu")=repnull(rsObj("GroupMenu"))
			Session("GroupSort")=repnull(rsObj("GroupSort"))
			'查找
			dim LanguageID
			LanguageID=rsObj("LanguageID")

			if isnul(LanguageID) then 				
				set rs=conn.exec("select LanguageID,LanguagePath,Alias from {prefix}Language where IsDefault=1","exe")
				if rs.eof then					
					echoErr err_15,15,"没有设置默认语言"
				else
					LanguageID=rs(0)
				end if
			else 	
				set rs=conn.exec("select LanguageID,LanguagePath,Alias from {prefix}Language where LanguageID="&LanguageID,"exe")				
				LanguageID=rs(0)	
			end if 
			randomize
			dim randnum
			randnum=clng(rnd*99999999)
			Conn.Exec"update {prefix}User set adminrand='"& randnum &"' where UserID="&rsObj("UserID"),"exe"	
			Session("adminrand")=rsObj("adminrand")	
			Session("languageID")=LanguageID			
			Session("LanguagePath")=rs("LanguagePath")
			Session("LanguageAlias")=rs("Alias")
			

			rs.close : set rs=nothing
			Conn.Exec"update {prefix}User set LastLoginTime='"&now()&"',LastLoginIP='"&getIp()&"',LoginCount=LoginCount+1 where UserID="&rsObj("UserID"),"exe"	
			if atype=0 then		
			response.Redirect("index.asp")
			else
			response.Redirect("index_user.asp")	
			end if
		end if
	else
		alertMsgAndGo "用户名或密码错误!","-1"
	end if
	rsObj.Close() : set rsObj=Nothing	
elseif action = "logout" then
	Session("adminName")=""
	Session("adminId")=""
	Session("groupMenu")=""
	Session("SceneMenu")=""
	response.Redirect("login.asp"):response.End()
elseif action = "relog" then
	alertMsgAndGo "对不起,您的登录状态已经失效,请重新登录!","login.asp"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><%=setting.siteTitle%>-开源企业网站管理系统- Powered by AspCms2.0</title>
<style type="text/css">
<!--
*{
	padding:0px;
	margin:0px;
	font-family:Verdana, Arial, Helvetica, sans-serif;
}
body {
	margin: 0px;
	font-size:12px;
}
input{
	vertical-align:middle;
}
img{
	border:none;
	vertical-align:middle;
}
a{
	color:#2366A8;
	text-decoration: none;
}
a:hover{
	color:#2366A8;
	text-decoration:underline;
}
.main{
	width:640px;
	margin:120px auto 0px;
	background:#FFF;
	padding-bottom:10px;
}

.main .title{
	width:600px;
	height:50px;
	margin:0px auto;
	background:url(images/login_toptitle.jpg) -10px 0px no-repeat;
	text-indent:326px;
	line-height:46px;
	font-size:14px;
	letter-spacing:2px;
	color:#F60;
	font-weight:bold;
}

.main .login{
	width:630px;
	margin:20px auto 0px;
	overflow:hidden;
}
.main .login .inputbox{
	width:230px;
	float:right;
	background-image: url(images/login_input_hr.gif);
	background-repeat: no-repeat;
	background-position: left center;
	padding-left: 20px;
}
.main .login .inputbox dl{
	height:28px;
	clear:both;
}
.main .login .inputbox dl dt{
	float:left;
	width:50px;
	height:25px;
	line-height:25px;
	text-align:right;
	padding-right: 5px;
}
.main .login .inputbox dl dd{
	width:160px;
	float:left;
	padding-top:1px;
}
.main .login .inputbox dl dd input{
	font-size:12px;
	border:1px solid #dcdcdc;
	height: 18px;
	line-height: 18px;
	padding-right: 2px;
	padding-left: 2px;
	background-image: url(images/login_input_bg.gif);
	background-repeat: no-repeat;
	background-position: left top;
}
.main   .login   .inputbox   dl   .input  {
	width:57px;
	height:23px;
	background:url(images/login_submit.gif) no-repeat;
	border:none;
	cursor:pointer;
}


.main .login .butbox{
	float:left;
	width:350px;
}
.main .login .butbox dl{
}
.main .login .butbox dl dt{
	height:50px;
	padding-top:5px;
	background-image: url(images/logo_01.gif);
	background-repeat: no-repeat;
}
.main .login .butbox dl dd{
	height:21px;
	line-height:21px;
}



.main .msg{
	clear:both;
	line-height:20px;
	border:1px solid #DCDCDC;
	color:#888;
	margin-left: 55px;
	background-color: #FFFFFF;
	padding-right: 2px;
	padding-left: 2px;
	width: 130px;
	margin-bottom: 3px;
}

.copyright{
	width:640px;
	text-align:center;
	font-size:10px;
	color:#999999;
	margin-top: 50px;
	margin-right: auto;
	margin-bottom: 10px;
	margin-left: auto;
}
.copyright a{
	font-weight:bold;
	color:#F63;
	text-decoration:none;
}
.copyright a:hover{
	color:#000;
}
-->
</style>
<script type="text/javascript" language="javascript">
<!--
	window.onload = function (){
		var txtUserName = document.getElementById("username");
		var txtPassword = document.getElementById("password");
		changeimg();
		var username = readCookie("username");
		if(username!="")
		{
		txtUserName.value = username;
		txtPassword.focus();
		}
		else	txtUserName.focus();
		document.getElementById("code").value = "";
	}
function changeimg(){document.getElementById('SeedImg').src='../inc/checkcode.asp?'+Math.random();}
function writeCookie(name, value, hours){
  var expire = "";
  if(hours != null)  {
    expire = new Date((new Date()).getTime() + hours * 3600000);
    expire = "; expires=" + expire.toGMTString();
  }
  document.cookie = name + "=" + escape(value) + expire;

}
function readCookie(name){
  var cookieValue = "";
  var search = name + "=";
  if(document.cookie.length > 0) { 
    offset = document.cookie.indexOf(search);
    if (offset != -1)  { 
      offset += search.length;
      end = document.cookie.indexOf(";", offset);
      if (end == -1) end = document.cookie.length;
      cookieValue = unescape(document.cookie.substring(offset, end))
    }
  }
  return cookieValue;
}


-->
</script>
</head>
<body>

	<div class="main">
		<div class="login">
		<form method="post" action="?action=login" onsubmit="writeCookie('username',this.username.value, 1)">
            <input type="hidden" name="gotopage" value="/aspcms/index.asp">
            <input type="hidden" name="dopost" value="login">
            <div class="inputbox">
				<dl>
					<dt>用户名：</dt>
					<dd><input type="text"  id="username" name="username" size="20" maxlength="20" onfocus="this.style.borderColor='#239fe3'" onblur="this.style.borderColor='#dcdcdc'" />
					</dd>
				</dl>
				<dl>
					<dt>密码：</dt>
					<dd><input type="password" id="password" maxlength="20" name="password"  size="20" onfocus="this.style.borderColor='#239fe3'" onblur="this.style.borderColor='#dcdcdc'" />
					</dd>
				</dl>
                <%if admincode=1 then%>
				<dl>
					<dt>验证码：</dt>
					<dd><input type="text" id="code" name="code"   size="10" maxlength="4" onfocus="this.style.borderColor='#239fe3'" onblur="this.style.borderColor='#dcdcdc'" /><img src="../inc/checkcode.asp" id="SeedImg" align="absmiddle" style="cursor:pointer;" border="0" alt="点我刷新" title="点我刷新" onclick="changeimg()" />
					 
					</dd>
				</dl>
                <%end if%>
                <dl>
					<dt>类型：</dt>
					<dd><input name="atype" type="radio" value="1" checked="checked"/>简单版  <input name="atype" type="radio" value="0" />全功能版
					</dd>
				</dl>
                
                	
				              		
		        <dl>
				  <dt>&nbsp;</dt>
					<dd>
					  <input name="submit" type="submit" value="" class="input" />
		  </dd>
				</dl>
            </div>
            <div class="butbox">
            <dl>
					<dt></dt>
			  <dd>
                    <span style="color:#505050;font-size:12px;display:block">ASPCMS!是 <a href="http://www.4i.com.cn"  target="_blank" style="color:#315D80">上谷网络</a> 旗下 <a href="http://www.aspcms.com" target="_blank" style="color:#315D80">开源创新实验室</a> 推出的以企业建站为基础的专业建站系统，帮助其实现一站式服务。</span>
                   
               <a href="/">返回首页</a><!-- | <a href="http://www.4i.com.cn">上谷网络</a> | <a href="http://www.aspcms.com">开源创新实验室</a>			   --></dd>
			  </dl>
			</div>
		</form>
        <div style="clear:both"></div>
	  </div>
	</div>
	
	<div class="copyright">
		Powered by <a href="http://www.aspcms.com">www.aspcms.com</a> Copyright &copy;2010-2013 
	</div>

</body>
</html>
