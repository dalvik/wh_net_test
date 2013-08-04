<!--#include file="vote_SettingFun.asp" -->

<%
dim votenames,votenums,i,votetitle,votename,votenum,votetype,sql
dim kvArr

sql = "select * from {prefix}vote "
		kvArr = Conn.Exec(sql,"arr")

if getDataCount("select Count(*) from {prefix}vote ") >0 then
votetitle = kvArr(1,0)
votename = kvArr(2,0)
votenum = kvArr(3,0)
votetype = kvArr(4,0)
end if

votenames=split(votename,"|")
echo "document.write(""<form action=\""/plug/vote/vote.asp\"" method=\""post\"" target=\""_blank\"">"");"
echo "document.write(""<table width=\""100%\"" border=\""0\"">"");"
echo "document.write(""<tr>"");"
echo "document.write(""<td>"& votetitle &"</td>"");"
echo "document.write(""</tr>"");"
for i=0 to ubound(votenames)
	echo "document.write(""<tr>"");"
	if votetype=0 then
	echo "document.write(""<td><input name=\""selectid\"" type=\""radio\"" value=\"""& i &"\"" />"& votenames(i) &"</td>"");"
	else
	echo "document.write(""<td><input name=\""selectid\"" type=\""checkbox\"" value=\"""& i &"\"" />"& votenames(i) &"</td>"");"
	end if
	echo "document.write(""</tr>"");"
next
echo "document.write(""<tr>"");"
echo "document.write(""<td><input name=\""submit\"" type=\""submit\"" value=\""提交\""/>&nbsp;<input name=\""res\"" type=\""button\"" value=\""查看\"" onclick=\""window.open('/plug/vote/r.asp',target='_blank')\""/></td>"");"
echo "document.write(""</tr>"");"
echo "document.write(""</table>"");"

echo "document.write(""</form>"");"
%>

