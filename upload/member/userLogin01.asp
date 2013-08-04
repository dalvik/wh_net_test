<!--#include file="../inc/AspCms_SettingClass.asp" -->
document.write("<div id=\"index_prolist\">");
<%
if  isnul(session("loginstatus")) or session("loginstatus")="0" then	
%>
document.write("      <form method=\"post\" name=\"\" action=\"<%=sitepath%>/member/login.asp?action=login\">");
document.write("    <table border=\"0\" cellpadding=\"2\" cellspacing=\"0\" align=\"left\">");
document.write("        <tr>");
document.write("          <td height=\"17\">用户名：</td>");
document.write("          <td><input type=\"text\" name=\"LoginName\"  style=\"BORDER: #CCCCCC 1px solid;WIDTH: 70px; height:15px;\" size=16 maxlength=\"50\" onfocus=\"this.select();\" color:#0099ff /></td>");
document.write("          <td height=\"17\">密  码：</td>");
document.write("          <td><input type=\"password\" name=\"userPass\"  style=\"BORDER: #CCCCCC 1px solid;WIDTH: 70px; height:15px;\" size=16 maxlength=\"50\" onfocus=\"this.select();\" color:#0099ff /></td>");
document.write("          <td height=\"17\">验证码：</td>");
document.write("          <td><input type=\"text\" name=\"code\"  style=\"BORDER: #CCCCCC 1px solid;WIDTH: 40px; height:15px;\" size=10 maxlength=\"4\" onfocus=\"this.select();\" color:#0099ff /><a onClick=\"SeedImg.src=\'<%=sitepath%>/inc/checkcode.asp\'\"><img src=\"<%=sitepath%>/inc/checkcode.asp\" id=\"SeedImg\" align=\"absmiddle\" style=\"cursor:pointer;\" border=\"0\" /></a></td>");
document.write("          <td height=\"18\" ></td>");
document.write("          <td><input type=\"submit\" value=\"登陆\" style=\"BORDER:none;color:#0099ff;background:none;cursor:pointer;width:35px; text-align:center;padding:0px;\" />&nbsp;");
document.write("              <input type=\"button\" value=\"注册\" style=\"BORDER:none;color:#0099ff;background:none;cursor:pointer;width:35px; text-align:center;padding:0px;\" onclick=\"location.href=\'<%=sitepath%>/member/reg.asp\'\" /></td>");
document.write("        </tr>");
document.write("    </table>");
document.write("      </form>");
<%	else%>
document.write("      <table border=\"0\" cellpadding=\"2\" cellspacing=\"0\" align=\"left\">");
document.write("        <tr>");
document.write("          <td height=\"28\">用户名：</td>");
document.write("          <td><font color=\"red\"><%=session("loginName")%></font></td>");

document.write("          <td height=\"28\"><a href=\"<%=sitepath%>/member/userinfo.asp\"><font color=\"#2485c9\">用户面板</font></a></td>");
document.write("          <td><a href=\"<%=sitepath%>/member/login.asp?action=logout\"><font color=\"#2485c9\">退出登录</font></a></td>");
document.write("        </tr>");
document.write("        </table>");
<%
end if
%>
document.write("<DIV id=index_search3></DIV></div>");