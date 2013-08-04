<!--#include file="AspCms_53kfFun.asp" -->
<%getComplanySetting%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
a:link {
	color: #000099;
}
a:hover {
	color: #FF0000;
}
-->
</style>
</HEAD>
<BODY>


<DIV class=formzone>
<DIV class=namezone>53KF设置</DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<%if service53kf=0 then%>
<FORM name="form" action="?action=reg" method="post" >
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>    
  <TR>
    <TD height="28" colspan="2" class=innerbiaoti>一键安装53KF<div style="display:none;"><iframe src=" http://www.53kf.com?yx_from=chancoo" frameborder="0"></iframe></div></TD>
  </TR>
  <TR>
    <TD align=middle height=30><input type="submit" name="button" id="button" value="一键注册安装53KF"></TD>
    <TD>&nbsp;</TD>
  </TR>

    </TBODY>
</TABLE>
</FORM>
<%else%>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>    
  <TR>
    <TD height="28" colspan="2" class=innerbiaoti>53KF帐号信息</TD>
  </TR>
  <TR>
    <TD height=30 colspan="2" align=left><strong>恭喜您！您的53KF用户帐号已在53KF官网申请成功！可前往53KF官网下载<a href="http://www.53kf.com/download.php" target="_blank"><span class="STYLE1">53KF登录器</span></a>进行管理。</strong></TD>
    </TR>
  <TR>
    <TD width="18%" height=30 align=left>53KF使用状态</TD>
    <TD width="82%" align="left"><form name="form1" method="post" action="?action=Status">
      <label>
        <input type="radio" name="kfStatus" id="kfStatus" value="1" <%if service53kfStatus=1 then%> checked<%end if%>>开启
         <input type="radio" name="kfStatus" id="kfStatus" value="0" <%if service53kfStatus=0 then%> checked<%end if%>>关闭        </label>&nbsp;&nbsp;&nbsp;&nbsp;<input name="" type="submit" value="保存">
    </form>    </TD>
  </TR>
  <TR>
    <TD align=left height=30>客服帐号</TD>
    <TD align="left"><%=service53kfaccount%></TD>
  </TR>
  <TR>
    <TD align=left height=30>客服工号</TD>
    <TD align="left"><%=service53workid%></TD>
  </TR>
  <TR>						
    <TD align=left height=30>登录密码</TD>
    <TD align="left"><%=service53kfpasswd%></TD>
  </TR>
    </TBODY>
</TABLE>
<%end if%>
</DIV>
</DIV>


</BODY></HTML>

