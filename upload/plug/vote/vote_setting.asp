<!--#include file="vote_SettingFun.asp" -->
<!--#include file="vote_config.asp" -->
<%CheckLogin()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK href="style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.3790.4807" name=GENERATOR>
</HEAD>
<BODY>
<DIV class=formzone>
<FORM name="" action="?action=save" method="post">
<DIV class=tablezone>

<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
<TBODY>
    <TR>
    <TD width="356" class=innerbiaoti><STRONG>投票选项设置</STRONG></TD>
        <TD class=innerbiaoti width=507 height=28></TD>
    </TR>
    <TR class=list>
      <TD>投票标题设置</TD>
      <TD><INPUT class="input" size="30" style="width:80%" value="<%=votetitle%>" name="votetitle"/></TD>
    </TR>
    <TR class=list>
        <TD>投票选项设置（请用|隔开,例如 a|b|c|d）</TD>
        <TD width=507><INPUT class="input" size="30" style="width:80%" value="<%=votename%>" name="votename"/></TD>
    </TR>
   
    <TR class=list>
        <TD>投票票数设置（请与选项数量一致，按以上例子为 0|2|3|4）</TD>
        <TD width=507><INPUT class="input" size="30" style="width:80%" value="<%=votenum%>" name="votenum"/></TD>
    </TR>
     
     <TR class=list>
      <TD>投票类型</TD>
      <TD><label>
        <select name="votetype" id="votetype">
        <option value="0" <%if votetype=0 then%> selected<%end if%>>单选</option>
        <option value="1"<%if votetype=1 then%> selected<%end if%>>多选</option>
        </select>
      </label></TD>
    </TR><TR class=list>
       <TD>调用JS</TD>
       <TD>&lt;script src=&quot;{aspcms:sitepath}/plug/vote/index.asp&quot; language=&quot;JavaScript&quot;&gt;&lt;/script&gt;<br></TD>
     </TR>
</TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="保存" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 
 </DIV></FORM>
</DIV>
</BODY>
</HTML>
