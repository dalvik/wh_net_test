<%
Const runMode=0	'��վ����ģʽ��0Ϊ��̬��1Ϊ��̬��
Const sitePath=""	'��վ��Ŀ¼ ����:/cms
Const accessFilePath="data/#data.asp"	'access���ݿ��ļ�·��
Const dbType=0  '���ݿ����ͣ�0Ϊaccess��1Ϊsqlserver��
Const databaseServer="."  'sqlserver���ݿ��ַ
Const databaseName="aspcms"  'sqlserver���ݿ�����
Const databaseUser="sa"  'sqlserver���ݿ��˺�
Const databasepwd="sa"  'sqlserver���ݿ�����
Const fileExt=".html"	'�����ļ���չ����htm,html,asp��	
Const upLoadPath="upLoad"	'�ϴ��ļ�Ŀ¼
Const textFilter=""	'���ֹ���
Const tablePrefix="AspCms_"	'���ݿ�ǰ׺
Const upFileSize=20000	'�ϴ��ļ���С����KB
Const upFileWay=1	'�ϴ��ļ���ʽ����(1,������ϴ�,)
Const htmlDir="aspcms"	'�ĵ�HTMLĬ�ϱ���·��

'������
Const siteMode=1	'��վ״̬��0Ϊ�رգ�1Ϊ���У�
Const siteHelp="����վ����������ر���"	'��վ״̬��0Ϊ�رգ�1Ϊ���У�
Const admincode=1  '��̨��¼��֤�루0Ϊ�رգ�1������
Const switchFaq=0	'���Կ��أ�0Ϊ�رգ�1Ϊ������
Const switchFaqStatus=0 '������˿��أ�0Ϊ�����ã�1Ϊ���ã�
Const switchComments=1	'���ۿ��أ�0Ϊ�رգ�1Ϊ������
Const switchCommentsStatus=0	'��������Ƿ����ã�0Ϊ�����ã�1Ϊ���ã�


Const waterMark=1	'ˮӡ(0,1)
Const waterType=0	'ˮӡ����(0Ϊ���֣�1ͼƬ)
Const waterMarkFont="ˮӡʾ��"	'ˮӡ����
Const waterMarkPic="/upLoad/other/month_1212/201212232322181539.jpg"	'ˮӡͼƬ
Const markPicWidth=80
Const markPicHeight=25
Const markPicAlpha =0.5
Const waterMarkLocation="5"	'λ��

'�ʼ����� 
Const smtp_usermail="aspcmstest@163.com"	'�ʼ���ַ
Const smtp_user="aspcmstest"	'�ʼ��˺� 
Const smtp_password="aspcms.cn"	'�ʼ����� 
Const smtp_server="smtp.163.com"	'�ʼ�������

'���ѹ���
Const messageAlertsEmail="513518064@qq.com"	'�ʼ���ַ
Const messageReminded=1	'��������
Const orderReminded=1	'��������
Const commentReminded=1	'��������
Const applyReminded=1	'ӦƸ����

Const sortTypes="��ƪ,����,��Ʒ,����,��Ƹ,���,����,��Ƶ"

 
Const dirtyStr="�ڳ�<br>����"

'���߿ͷ�
Const serviceStatus=1	'���߿ͷ���ʾ״̬
Const serviceStyle=1	'��ʽ
Const serviceLocation="left"	'��ʾλ��
Const serviceQQ="����֧��|506232687 ��Ʒ��ѯ|8887443" 'QQ
'Const serviceEmail="234324324"	'����
'Const servicePhone="43244324324"	'�绰
Const serviceWangWang="����һ��|123456 ����2��|8887443"	'����
Const serviceContact="/about/?19.html"	'��ϵ��ʽ����
Const servicekfStatus=1	
Const servicekf=""	'

'53�ͷ�
Const service53kfStatus=0	'53KF��ʾ״̬
Const service53kf=0	'53KF����״̬
Const service53kfaccount="" '53KF�ʺ�
Const service53workid="" '53KF����
Const service53kfpasswd="" '53KF����


'�õ�Ƭ����a
Const slidestyle=0	'�õ�Ƭ��ʽ
Const slideImgs="/upLoad/slide/month_1109/201109301841433233.jpg, /upLoad/slide/month_1109/201109301841503128.jpg, /upLoad/slide/month_1109/201109301842008590.jpg,"	'ͼƬ��ַ
Const slideLinks="http://www.4i.com.cn, http://www.4i.com.cn, http://www.4i.com.cn,"	'���ӵ�ַ
Const slideTexts="ASPCMS��ҵ��վ��վƽ̨--�õ�Ƭ1, ASPCMS��ҵ��վ��վƽ̨--�õ�Ƭ2, ASPCMS��ҵ��վ��վƽ̨--�õ�Ƭ3,"	'����˵��
Const slideWidth="1001"	'��
Const slideHeight="244"	'��
Const slideTextStatus=1	'����˵������
Const slideNum=3	'����˵������

'�õ�Ƭ����B
Const slidestyleB=0	'�õ�Ƭ��ʽ
Const slideImgsB="/upLoad/slide/month_1109/201109301841433233.jpg, /upLoad/slide/month_1109/201109301841503128.jpg, /upLoad/slide/month_1109/201109301842008590.jpg,"	'ͼƬ��ַ
Const slideLinksB=", ,,"	'���ӵ�ַ
Const slideTextsB=", ,,"	'����˵��
Const slideWidthB="960"	'��
Const slideHeightB="263"	'��
Const slideTextStatusB=0	'����˵������
Const slideNumB=3	'����˵������

'�õ�Ƭ����C
Const slidestyleC=0	'�õ�Ƭ��ʽ
Const slideImgsC=","	'ͼƬ��ַ
Const slideLinksC=","	'���ӵ�ַ
Const slideTextsC=","	'����˵��
Const slideWidthC="202"	'��
Const slideHeightC="202"	'��
Const slideTextStatusC=1	'����˵������
Const slideNumC=1	'����˵������

'�õ�Ƭ����D
Const slidestyleD=1	'�õ�Ƭ��ʽ
Const slideImgsD=","	'ͼƬ��ַ
Const slideLinksD=","	'���ӵ�ַ
Const slideTextsD=","	'����˵��
Const slideWidthD="203"	'��
Const slideHeightD="203"	'��
Const slideTextStatusD=1	'����˵������
Const slideNumD=1	'����˵������


Const GoogleAPIKey=""
Const GoogleMapLat=30.593086
Const GoogleMapLng=114.30536



Const dirtyStrToggle=1

%>