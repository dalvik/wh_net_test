<!--#include file="../inc/AspCms_SettingClass.asp" -->
document.write("<div id=\"index_prolist\">");

document.write("<DIV id=index_search2>");
document.write("  <div class=query>");
<%
if  isnul(session("loginstatus")) or session("loginstatus")="0" then	
%>
document.write("      <form method=\"post\" name=\"\" action=\"<%=sitepath%>/member/login.asp?action=login\">");
document.write("    <table border=\"0\" cellpadding=\"2\" cellspacing=\"0\" align=\"center\">");
document.write("        <tr>");
document.write("          <td height=\"24\">�û�����</td>");
document.write("          <td><input type=\"text\" name=\"LoginName\"  style=\"BORDER: #CCCCCC 1px solid;WIDTH: 120px; height:18px;\" size=16 maxlength=\"50\" onfocus=\"this.select();\" color:#0099ff /></td>");
document.write("        </tr>");
document.write("        <tr>");
document.write("          <td height=\"24\">��  �룺</td>");
document.write("          <td><input type=\"password\" name=\"userPass\"  style=\"BORDER: #CCCCCC 1px solid;WIDTH: 120px; height:18px;\" size=16 maxlength=\"50\" onfocus=\"this.select();\" color:#0099ff /></td>");
document.write("        </tr>");
document.write("        <tr>");
document.write("        <td height=\"24\" align=\"left\">��֤�룺</td>");
document.write("        <td align=\"left\"><input type=\"text\" name=\"code\"  style=\"BORDER: #CCCCCC 1px solid;WIDTH: 50px; height:18px;\" size=10 maxlength=\"4\" onfocus=\"this.select();\" color:#0099ff /><a onClick=\"SeedImg.src=\'<%=sitepath%>/inc/checkcode.asp\'\"><img src=\"<%=sitepath%>/inc/checkcode.asp\" id=\"SeedImg\" align=\"absmiddle\" style=\"cursor:pointer;\" border=\"0\" /></a></td>");
document.write("        </tr>");
document.write("        <tr>");
document.write("          <td height=\"18\" colspan=\"2\"><input type=\"submit\" value=\"��½\" style=\"BORDER: #CCCCCC 1px solid;WIDTH: 60px; height:22px; cursor:pointer\"/>&nbsp;");
document.write("              <input type=\"button\" value=\"ע��\" onclick=\"location.href=\'<%=sitepath%>/member/reg.asp\'\" style=\"BORDER: #CCCCCC 1px solid;WIDTH: 60px; height:22px;cursor:pointer\"/></td>");

document.write("        </tr>");
document.write("    </table>");
document.write("      </form>");
<%	else%>
document.write("      <table border=\"0\" cellpadding=\"2\" cellspacing=\"0\" align=\"center\">");
document.write("        <tr>");
document.write("          <td height=\"28\">�û�����</td>");
document.write("          <td><font color=\"red\"><%=session("loginName")%></font></td>");
document.write("        </tr>");
document.write("        <tr>");
document.write("        <tr>");
document.write("          <td height=\"28\"><a href=\"<%=sitepath%>/member/userinfo.asp\"><font color=\"#2485c9\">�û����</font></a></td>");
document.write("          <td><a href=\"<%=sitepath%>/member/login.asp?action=logout\"><font color=\"#2485c9\">�˳���¼</font></a></td>");
document.write("        </tr>");
document.write("        </table>");
<%
end if
%>
document.write("  </div>");
document.write("</DIV>");
document.write("<DIV id=index_search3></DIV></DIV>");