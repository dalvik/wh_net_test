<%
'阿里云开放平台
Const OSS_ACCESS_ID="8gkbkx8a57cl0x65i3oa546p"	'设置ACCESS_ID
Const OSS_ACCESS_KEY="PrXnKlPvImrbVGXjQ6wzCkYxniM=" '设置ACCESS_KEY

Const DEFAULT_OSS_HOST="storage.aliyun.com"   '阿里云OSS服务公网地址

'Const DEBUG=0        '是否输出DEBUG（0->不输出，1->输出）
Const LANG="gbk"        '语言类型 
Const OSS_CONTENT_MD5="aspcms0306"   'header头部MD5加密
Const OSS_OBJECT_MAXLEN=100    '对象列表显示条数
Const OSS_OBJECT_EXPIRE=3600   '设置签名外链的过期时间为3600秒
Const OSS_OBJECT_EXT=".mdb"   '默认备份文件的扩展名  

'Const FILE_PATH=""  '备份文件目录
Const PAST_BACKUP="2012-5-26"    '上次备份时间
Const DEFAULT_ACL="private"   '默认创建的bucket的权限
Const BUCKET_PREFIX="aspcms"   'bucket前缀，不能为空，且只能是英文字母或数字，不能超过10个字符

'Const MAX_UPLOAD_FILE_SIZE=128*1024*1024         '上传文件的最大值,默认值128M
Const MAX_EXECUTE_TIME="3600"          '

Const OSS_NAME="OSS_ASP_云存储" 
Const OSS_VERSION="1.0"
Const OSS_AUTHOR="yao_62995@163.com"     'Author

'------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------

%>