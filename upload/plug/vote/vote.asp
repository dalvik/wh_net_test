<!--#include file="vote_SettingFun.asp" -->
<!--#include file="vote_config.asp" -->
<%
dim i,j,newvotenum,newvotenums
dim selectid:selectid=getForm("selectid","post")
if selectid="" then
	echo "<script>alert('未选择投票选项');window.close();</script>"
	die
end if
if request.Cookies("vote")="yes" then
	echo "<script>alert('已投过票');window.close();</script>"
	die
end if
dim selectids:selectids=split(selectid,",")
dim votenums:votenums=split(votenum,"|")

for i=0 to ubound(votenums)
	for j=0 to ubound(selectids)
		newvotenum=votenums(i)
		if i=cint(selectids(j)) then
			newvotenum=cint(votenums(i))+1
			exit for
			
		end if
	next
	if i=0 then
		newvotenums=newvotenum
	else
		newvotenums=newvotenums&"|"&newvotenum
	end if
next

dim templateobj,configStr,configPath
	configPath="vote_config.asp"
	set templateobj=new TemplateClass
	configStr=loadFile(configPath)
		configStr=templateobj.regExpReplace(configStr,"Const votenum=""(\S*?)""","Const votenum="""&newvotenums&"""")


	createTextFile configStr,configPath,""
	set templateobj=nothing
	response.Cookies("vote")="yes"
	Response.Cookies("vote").Expires = DateAdd("s",3600*24,Now())
	alertMsgAndGo "投票成功","/plug/vote/r.asp"
%>