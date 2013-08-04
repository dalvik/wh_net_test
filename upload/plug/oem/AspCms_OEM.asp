<!--#include file="AspCms_OEMFun.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE>OEM版权信息更新</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.3790.4807" name=GENERATOR>
<script language="javascript" type="text/javascript" src="../js/getdate/WdatePicker.js"></script>
<script language="javascript">


function getObject(objectId) {
 if(document.getElementById && document.getElementById(objectId)) {
 // W3C DOM
 return document.getElementById(objectId);
 }
 else if (document.all && document.all(objectId)) {
 // MSIE 4 DOM
 return document.all(objectId);
 }
 else if (document.layers && document.layers[objectId]) {
 // NN 4 DOM.. note: this won't find nested layers
 return document.layers[objectId];
 }
 else {
 return false;
 }
}

function showHide(objname){
    var obj = getObject(objname);
   
		obj.style.display = "block";
	
}
function disHide(objname){
    var obj = getObject(objname);
   
		obj.style.display = "none";
	
}
</script>
<style type="text/css">
<!--
.STYLEnote {color: #FF0000}
-->
</style>
</head>
<body>
<FORM id= name= action="?action=down" method=post>
<DIV class=formzone>
<DIV class=namezone>OEM版权信息更新</DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
 
  <TR>
    <TD align=center height=30>&nbsp;</TD>
    <TD>
      <strong>*<span class="STYLEnote">重要提示</span>:如果修改了后台目录，请进入/plug/oem/AspCms_OEMFun.asp 中将 adminpath="admin" 中的admin改为您的后台目录，例如如果您的后台目录改成了aspcms,文件中则为adminpath="aspcms"</strong></TD>
  </TR>
  <TR>						
    <TD align=center width=125 height=30>OEM用户名</TD>
    <TD width="1018"><input type="text" name="username" id="username">
      &nbsp;<a href="http://www.aspcms.com/aspcms-41029-1-1.html" target="_blank">没有账号?</a></TD>
  </TR>

  <TR>
<TD align=center width=125 height=30>OEM密码</TD>
    <TD width="1018"><input type="password" name="password" id="password"></TD>
  </tr>
  <TR>
    <TD align=center width=125 height=30>&nbsp;</TD>    
    <TD><label>
      <input type="submit" name="button" id="button" value="提交" onClick="showHide('loadpic');disHide('button')"><img id="loadpic" src="load.gif" style="display:none">
    </label></TD>
 </tr>
  <TR>
    <TD align=center height=30>&nbsp;</TD>
    <TD>&nbsp;</TD>
  </tr>
  <TR>
    <TD rowspan="11" align=center><img src="logo.jpg" ></TD>
    <TD height="30"><a href="http://www.aspcms.com/aspcms-41029-1-1.html" class="STYLEnote" target="_blank"><strong>如何oem？</strong></a></TD>
  </tr>
  <TR>
    <TD height="30">ASPCMS|OEM更换LOGO\新闻\服务信息步骤：</TD>
  </tr>
  <TR>
    <TD height="30">1、购买邀请码。</TD>
  </tr>

  <TR>
    <TD height="30">2、获得邀请码，申请OEM账号</TD>
  </tr>
    <TR>
    <TD height="30">3、填写DIY后台LOGO\新闻\服务信息。</TD>
  </tr>
    <TR>
      <TD height="30">4、ASPCMS安装OEM插件。</TD>
    </tr>
    <TR>
      <TD height="30">5、插件中直接填写用户名和密码。</TD>
    </tr>
    <TR>
      <TD height="30">6、点击提交。</TD>
    </tr>
    <TR>
      <TD height="30">7、注销账户，重新登录。</TD>
    </tr>
    <TR>
      <TD height="30">8、OEM成功。</TD>
    </tr>
    <TR>
    <TD height="30">ASPCMS|OEM帮助地址：<a href="http://www.aspcms.com/aspcms-40654-1-1.html" target="_blank">http://www.aspcms.com/aspcms-40654-1-1.html</a></TD>
  </tr>
   <TR>
    <TD height=30 colspan="2" align=left>&nbsp;</TD>
    </tr>
    </TBODY></TABLE>

</DIV>
</DIV>
</FORM>
</body>
</html>
