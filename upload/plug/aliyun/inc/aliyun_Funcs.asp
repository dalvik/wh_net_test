<!--#include file="HMAC-SHA1.asp" -->
<%
'----------------------
'获取bucket列表，返回数组
'----------------------
function bucket_list()
	Dim xmlhttp,xmlDOC,nodes,resultArr(),sign,gmttime,iterator
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("GET"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"/",OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "GET","http://"&DEFAULT_OSS_HOST&"/",false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
		xmlDOC.load(xmlhttp.responseXML)
		Set nodes=xmlDOC.selectNodes("//Buckets/Bucket")
		REDIM resultArr(nodes.length)
		for iterator=0 to nodes.length-1
			resultArr(iterator)=nodes(iterator).childnodes(0).text
		next
		bucket_list=resultArr
		'bucket_list=BytesToBstr(xmlhttp.responsebody,"gb2312")
		Set xmlDOC = nothing
	Else
		bucket_list=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function

'------------------------
'获得指定bucket的ACL
'------------------------
function get_bucket_acl(bucket)
	Dim xmlhttp,result,sign,gmttime
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("GET"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"text/xml"&Chr(10)+gmttime+Chr(10)&"/"&bucket&"?acl",OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "GET","http://"&DEFAULT_OSS_HOST&"/"&bucket&"?acl",false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","text/xml"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		result=split(BytesToBstr(xmlhttp.responsebody,"gb2312"),"<Grant>")
		get_bucket_acl=split(result(1),"</Grant>")(0)
	Else
		get_bucket_acl=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function
'------------------------
'设置指定bucket的ACL,目前只有三种acl private,public-read,public-read-write
'------------------------
function set_bucket_acl(bucket,acl)
	Dim xmlhttp,sign,gmttime
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("PUT"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"x-oss-acl:"&acl+Chr(10)&"/"&bucket,OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "PUT","http://"&DEFAULT_OSS_HOST&"/"&bucket,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "x-oss-acl",acl
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		set_bucket_acl="1"
	Else
		set_bucket_acl=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function


'------------------------
'创建bucket
'------------------------
function bucket_create(bucket,acl)
	Dim xmlhttp,sign,gmttime
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("PUT"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"x-oss-acl:"&acl+Chr(10)&"/"&bucket,OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "PUT","http://"&DEFAULT_OSS_HOST&"/"&bucket,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "x-oss-acl",acl
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		bucket_create="1"
	Else
		bucket_create=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function

'------------------------
'删除bucket
'------------------------
function bucket_delete(bucket)
	Dim xmlhttp,sign,gmttime
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("DELETE"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"/"&bucket,OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "DELETE","http://"&DEFAULT_OSS_HOST&"/"&bucket,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 204 Then
		bucket_delete="1"
	Else
		bucket_delete=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function


'------------------------
'删除object
'------------------------
function object_delete(bucket,object)
	Dim xmlhttp,sign,gmttime
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("DELETE"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"/"&bucket&"/"&object,OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "DELETE","http://"&DEFAULT_OSS_HOST&"/"&bucket&"/"&object,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 204 Then
		object_delete="1"
	Else
		object_delete=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function


'----------------------
'获取指定bucket下的object列表，返回数组
'----------------------
function object_list(bucket,prefix,length)
	Dim xmlhttp,xmlDOC,nodes,resultArr(),sign,gmttime,iterator
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("GET"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"/"&bucket,OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "GET","http://"&DEFAULT_OSS_HOST&"/"&bucket,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "delimiter","/"
	xmlhttp.setRequestHeader "prefix",prefix
	xmlhttp.setRequestHeader "max-keys",length
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
		xmlDOC.load(xmlhttp.responseXML)
		Set nodes=xmlDOC.selectNodes("//Contents")
		REDIM Preserve resultArr(nodes.length,3)
		for iterator=0 to nodes.length-1
			resultArr(iterator,0)=nodes(iterator).childnodes(0).text  '文件名
			resultArr(iterator,1)=nodes(iterator).childnodes(1).text  '时间
			resultArr(iterator,2)=nodes(iterator).childnodes(4).text  '大小
		next
		object_list=resultArr
		'bucket_list=BytesToBstr(xmlhttp.responsebody,"gb2312")
		Set xmlDOC = nothing
	Else
		object_list=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function

'------------------------
'获得某一个object的URL
'------------------------
function get_object_url(bucket,object)
	Dim xmlhttp,sign,gmttime
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("HEAD"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"/"&bucket&"/"&object,OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "HEAD","http://"&DEFAULT_OSS_HOST&"/"&bucket&"/"&object,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		get_object_url=xmlhttp.Status
	Else
		get_object_url=xmlhttp.Status
	End if
	Set xmlhttp = Nothing
end function

'------------------------
'获得某一个object的带签名的外链URL，当设置为私有权限时只能通过该地址下载object
'------------------------
function get_sign_url(bucket,object)
	Dim sign,timeline,url
	timeline=DateDiff("s","01/01/1970 00:00:00",DateAdd("h",-8,now))+OSS_OBJECT_EXPIRE
	sign=OTCMS_b64_hmac_sha1("GET"&Chr(10)+Chr(10)+Chr(10)&""&timeline&""&Chr(10)&"/"&bucket&"/"&object,OSS_ACCESS_KEY)
	url="http://"&DEFAULT_OSS_HOST&"/"&bucket&"/"&object&"?OSSAccessKeyId="&OSS_ACCESS_ID&"&Expires="&timeline&"&Signature="&url_encode(sign)
	get_sign_url=url
end function

'----------------------
'获得$bucket下的某个object,$object为文件，不能为目录
'----------------------
function get_object(bucket,object,filepath)
	Dim xmlhttp,xmlDOC,nodes,resultArr(),sign,gmttime,iterator
	'创建文件
	'Dim fso
	'Set fso = Server.CreateObject("Scripting.FileSystemObject")
	'if not fso.FileExists(filePath)  then
	'	fso.CreateTextFile filePath,false 
	'end if
	'写入文件流
	Set Ado=Server.CreateObject("Adodb.Stream") 'Ado对象
	filePath=Server.MapPath(filepath) '文件路径
	'Ado.Charset="gb2312"
	Ado.Mode = 3 '1 读，2 写，3 读写。
	Ado.Type = 1 '1 二进制，2 文本。
	Ado.Open
	Ado.LoadFromFile(filePath) '载入文件
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("GET"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"/"&bucket&"/"&object,OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "GET","http://"&DEFAULT_OSS_HOST&"/"&bucket&"/"&object,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		Ado.write(xmlhttp.responsebody)
		'Ado.SaveToFile filePath,1 
		get_object=BytesToBstr(xmlhttp.responsebody,"gb2312")
	Else
		get_object=BytesToBstr(xmlhttp.responsebody,"gb2312")
	End if
	Ado.Close '关闭文件对象
	Set xmlhttp = Nothing
end function

'----------------------
'创建文件夹(是虚拟文件夹)
'----------------------
function object_dir_create(bucket,object)
	Dim xmlhttp,sign,gmttime
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("PUT"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/x-www-form-urlencoded"&Chr(10)+gmttime+Chr(10)&"/"&bucket&"/"&object&"/",OSS_ACCESS_KEY)
	'response.Write(sign&"<br>")
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "PUT","http://"&DEFAULT_OSS_HOST&"/"&bucket&"/"&object&"/",false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/x-www-form-urlencoded"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		object_dir_create="1"
	Else
		object_dir_create=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function

'----------------------
'通过http body上传文件,适用于直接写入内容的上传，比较小的文件
'上传的文件被二进制编码，要用BytesToBstr反编译才能使用
'----------------------

function upload_file_by_content(bucket,object,filename)
	Dim xmlhttp,sign,gmttime,Ado,filePath,file_size
	'读取二进制文件流
	Set Ado=Server.CreateObject("Adodb.Stream") 'Ado对象
	filePath=Server.MapPath(filename) '文件路径
	echo filepath
	Ado.Charset="gb2312"
	Ado.Mode = 3 '1 读，2 写，3 读写。
	Ado.Type = 1 '1 二进制，2 文本。
	'Ado.Charset="utf-8"
	Ado.Open
	Ado.LoadFromFile(filePath) '载入文件
	file_size=int((Ado.size)/1000)
	if file_size>1000 then
		file_size=FormatNumber(file_size/1000,1)&"M"
	else
		file_size=file_size&"K"
	end if
	response.Write("上传文件大小为："&file_size&"bytes，正在上传中，请耐心稍候。。。")
	gmttime=GetGMT(now)
	'response.Write("<br>"&OTCMS_binb2b64(MD5("content",32)))
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("PUT"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"application/octet-stream"&Chr(10)+gmttime+Chr(10)&"/"&bucket&"/"&object,OSS_ACCESS_KEY)
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.setTimeouts 1000000,1000000,1000000,1000000
	xmlhttp.Open "PUT","http://"&DEFAULT_OSS_HOST&"/"&bucket&"/"&object,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","application/octet-stream"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "content-length",Ado.size
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	'response.Write(BytesToBstr(Ado.Read(),"gb2312"))
	xmlhttp.send(Ado.Read())
	If xmlhttp.Status=200 Then
		upload_file_by_content="1"
	Else
		upload_file_by_content=BytesToBstr(xmlhttp.responsebody,"gb2312")
	End if
	Ado.Close '关闭文件对象
	Set xmlhttp = Nothing
end function

'----------------------
'通过文件方式上传,适合小文件上传，大文件上传请使用multipart,其中object为完整文件名（包括路径，如“2011_01/temp.mdb”),未使用
'----------------------
function upload_file(bucket,object,filename)
	Dim xmlhttp,sign,gmttime,Ado,filePath
	'读取二进制文件流
	Set Ado=Server.CreateObject("Adodb.Stream") 'Ado对象
	filePath=Server.MapPath(filename) '文件路径
	Ado.Mode = 3 '1 读，2 写，3 读写。
	Ado.Type = 2 '1 二进制，2 文本。
	Ado.Open
	Ado.LoadFromFile(filePath) '载入文件
	response.Write(Ado.size)
	gmttime=GetGMT(now)
	sign="OSS "&OSS_ACCESS_ID&":"&OTCMS_b64_hmac_sha1("PUT"&Chr(10)+OSS_CONTENT_MD5+Chr(10)&"text/plain"&Chr(10)+gmttime+Chr(10)&"/"&bucket&"/"&object,OSS_ACCESS_KEY)
	Set xmlhttp = server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "PUT","http://"&DEFAULT_OSS_HOST&"/"&bucket&"/"&object,false
	xmlhttp.setRequestHeader "content-md5",OSS_CONTENT_MD5
	xmlhttp.setRequestHeader "content-type","text/plain"
	xmlhttp.setRequestHeader "date",gmttime
	xmlhttp.setRequestHeader "content-length",Ado.size
	xmlhttp.setRequestHeader "host",DEFAULT_OSS_HOST
	xmlhttp.setRequestHeader "authorization",sign
	xmlhttp.send()
	If xmlhttp.Status = 200 Then
		upload_file="1"
	Else
		upload_file=BytesToBstr(xmlhttp.responsebody,"gb2312")
	End if
	Ado.Close '关闭文件对象
	Set xmlhttp = Nothing
end function


'-----------------------
'转换编码
'-----------------------	
Function BytesToBstr(body,Cset)     
	dim   objstream     
	set   objstream   =   Server.CreateObject("adodb.stream")     
	objstream.Type   =   1     
	objstream.Mode   =3     
	objstream.Open     
	objstream.Write  body     
	objstream.Position   =   0     
	objstream.Type   =   2     
	objstream.Charset   =   Cset     
	BytesToBstr   =   objstream.ReadText       
	objstream.Close     
	set   objstream   =   nothing     
End Function     
'生成随机字符，n为字符的个数 ，该随机函数由大小写字母组成，不含数字
function MyRandc(n) 
dim thechr,i
thechr = "" 
for i=1 to n 
   dim zNum,zNum2 
   Randomize 
   zNum = cint(25*Rnd) 
   zNum2 = cint(10*Rnd) 
   if zNum2 mod 2 = 0 then 
    zNum = zNum + 97 
   else 
    zNum = zNum + 65 
   end if 
   thechr = thechr & chr(zNum) 
next 
MyRandc = thechr 
end function



'-----------------------
'获取GMT格式的时间
'-----------------------
Function GetGMT(dat)
    dim y,m,d,h,mm,s,r,week,mon,w
	dat=DateAdd("h",-8,dat)
    week=split("Sun,Mon,Tue,Wed,Thu,Fri,Sat",",")
    mon=split("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",",")
    y=year(dat)
    m=mon(month(dat)-1)
    d=day(dat):if d<10 then d="0"&d
    h=hour(dat):if h<10 then h="0"&h
    mm=minute(dat):if mm<10 then mm="0"&mm
    s=second(dat):if s<10 then s="0"&s
    w=week(weekday(dat)-1)
    GetGMT=w&", "&d&" "&m&" "&y&" "&h&":"&mm&":"&s&" GMT"
 End Function
 
 
'url编码，只替换默认的字符串
Function url_encode(ByVal urlstr)
	urlstr = Replace(urlstr, "+", "%2B")
	urlstr = Replace(urlstr, " ", "+")
	urlstr = Replace(urlstr, "=", "%3D")
	urlstr = Replace(urlstr, "&", "%26")
	urlstr = Replace(urlstr, ":", "%3A")
	urlstr = Replace(urlstr, "/", "%2F")
	url_encode = urlstr
End Function

'url解码，只替换默认的字符串
Function url_decode(ByVal urlstr)
	urlstr = Replace(urlstr, "%5F", "_")
	urlstr = Replace(urlstr, "%2E", ".")
	urlstr = Replace(urlstr, "%3D", "=")
	urlstr = Replace(urlstr, "%26", "&")
	urlstr = Replace(urlstr, "%3A", ":")
	urlstr = Replace(urlstr, "%2F","/")
	url_decode = urlstr
End Function
	


%>
