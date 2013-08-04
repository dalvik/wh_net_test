Dim db_conn,db_path
Set db_conn=Server.CreateObject("ADODB.Connection")
db_path=Server.MapPath(sitePath&"/"&accessFilePath)
db_conn.Open "provider=microsoft.jet.oledb.4.0;data source="&db_path

Dim smssort,smsRecord,smsnum
smssort=tablePrefix&"smssort"
smsRecord=tablePrefix&"smsRecord"
smsnum=tablePrefix&"smsnum"

'创建表{prefix}smssort
response.Write("正在创建表"&smssort&"...<br>")
if CheckTable(smssort)=0 then
	db_conn.execute("Create TABLE "&smssort&"(smssortid counter primary key,smssortname text(50))")
	response.Write("表"&smssort&"创建成功<br>")
else
	response.Write("表"&smssort&"已存在<br>")
end if

'创建表{prefix}smsRecord
response.Write("正在创建表"&smsRecord&"...<br>")
if CheckTable(smsRecord)=0 then
	db_conn.execute("Create TABLE "&smsRecord&"(smsrid counter primary key,smsrtime datetime,smsryes int default 0,smsrno int default 0,smsrcontent longText)")
	response.Write("表"&smsRecord&"创建成功<br>")
else 
	response.Write("表"&smsRecord&"已存在<br>")
end if

'创建表{prefix}smsnum
response.Write("正在创建表"&smsnum&"...<br>")
if CheckTable(smsnum)=0 then
	db_conn.execute("Create TABLE "&smsnum&"(smsid counter primary key,smsname text(50),smsnum text(50),smssortid int default 0,remark longText)")
	response.Write("表"&smsnum&"创建成功<br>")
else 
	response.Write("表"&smsnum&"已存在<br>")
end if

set db_conn=nothing
'--------------------
'判断数据库中是否存在指定表
'--------------------
Function CheckTable(myTable)
'列出数据库中的所有表
dim rs,getTableName
Dim db_conn,db_path
Set db_conn=Server.CreateObject("ADODB.Connection")
db_path=Server.MapPath(sitePath&"/"&accessFilePath)
db_conn.Open "provider=microsoft.jet.oledb.4.0;data source="&db_path
set rs=db_conn.openSchema(20) 
rs.movefirst 
	do Until rs.EOF 
	if rs("TABLE_TYPE")="TABLE" then 
	   'response.write rs("TABLE_NAME") & ","
	   getTableName=getTableName+rs("TABLE_NAME")&","
	end if 
	rs.movenext
	loop
rs.close
set rs=nothing
set db_conn=nothing
'判断数据库中是否存在此表（下面两行代码使待比较表和指定表前后都加上豆号，以精确比较）
dim getTableName2,myTable2
getTableName2=","&getTableName '此字符串后面已经有豆号
myTable2=","&myTable&","
If instr(getTableName2,myTable2)<>0 Then
   CheckTable=1 '存在
else
   CheckTable=0 '不存在
end if
End Function