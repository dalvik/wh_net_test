<!--#include file="inc/aliyun_config.asp" -->
<!--#include file="inc/aliyun_funcs.asp" -->
<%

'��ȡbucket�б���������
'Dim arr,i
'arr=bucket_list()
'for i=0 to UBound(arr)-1
'	response.Write(arr(i))
'next

'���ָ��bucket��ACL
'response.Write("<br>get_bucket_acl:"&get_bucket_acl("aspcms")&"<br>")

'����ָ��bucket��ACL,Ŀǰֻ������acl private,public-read,public-read-write
'if set_bucket_acl("aspcms","private")="1" then
' 	response.Write("set ok")
'else
'	response.Write("set false")
'end if

'����bucket
'if bucket_create("aspcms0","private")="1" then
' 	response.Write("create ok")
'else
'	response.Write("create false")
'end if

'ɾ��bucket
'if bucket_delete("aspcms0")="1" then
' 	response.Write("delete ok")
'else
'	response.Write("delete false")
'end if



'��ȡobject�б���������
'Dim objectarr,objecti,j
'objectarr=object_list("aspcms_2011","2011_10_",OSS_OBJECT_MAXLEN)
'for objecti=0 to UBound(objectarr)-1
	
'	response.Write(objectarr(objecti,0)&"***")
'	objectarr(objecti,1)=split(objectarr(objecti,1),"T")(0)
'	response.Write(objectarr(objecti,1)&"***")
'	response.Write(int(objectarr(objecti,2)/1000)&"***")
'	response.Write("<br>")
'next

'��ȡָ����object

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''������
'response.Write(get_object("aspcms_2011","2011_10_12.mdb","t.txt"))


'��ȡojbect��url
'response.Write(get_object_url("aspcms_2011","2011_10_12.mdb"))

'��ȡojbect�Ĵ�ǩ��url
response.Write(get_sign_url("aspcms_2011","2011_12_15.mdb",3600))


'�����ļ���
'if object_dir_create("aspcms","2050")="1" then
' 	response.Write("create dir ok")
'else
'	response.Write("create dir false")
'end if

'Dim txt
'txt=upload_file_by_content("aspcms_2012","2012_5_12.mdb","temp.txt")
'if txt="1" then
'	response.Write("create file ok")
'else
'	response.Write(txt)
'end if
%>
