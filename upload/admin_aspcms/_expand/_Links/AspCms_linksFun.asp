<!--#include file="../../../inc/AspCms_MainClass.asp" -->
<%
CheckAdmin("AspCms_Links.asp")

dim action : action=getForm("action", "get")
dim LinkType,LinkText,LinkDesc,ImageURL,LinkOrder,LinkStatus,LinkURL,ID,Linkgroup


Select case action
	case "del" : delLinks
	case "add" : addLinks
	case "edit" : editLinks
	case "on" : onOff "on", "Links", "LinkID", "LinkStatus", "", getPageName()
	case "off" : onOff "off", "Links", "LinkID", "LinkStatus", "", getPageName()
	case "ugroup" : updategroup	
	
End Select

Sub getContent
	ID = getForm("id","get")
	Dim rsObj 	: Set rsObj=Conn.Exec("select LinkType,LinkText,LinkDesc,ImageURL,LinkOrder,LinkStatus,LinkURL,Linkgroup from {prefix}Links where LinkID="&ID,"r1")
	if isnul(ID) then alertMsgAndGo "ID����Ϊ��","-1"	
	LinkType=rsObj(0)
	LinkText=rsObj(1)
	LinkDesc=rsObj(2)
	ImageURL=rsObj(3)
	LinkOrder=rsObj(4)
	LinkStatus=rsObj(5)
	LinkURL=rsObj(6)
	Linkgroup=rsObj(7)
	rsObj.close	:	Set rsObj = nothing
End Sub

Sub editLinks
	ID=getForm("LinkID","post")
	LinkType=getForm("LinkType","post")	
	LinkText=getForm("LinkText","post")	
	LinkDesc=getForm("LinkDesc","post")	
	ImageURL =getForm("ImageURL","post")	
	LinkOrder=getForm("LinkOrder","post")
	LinkStatus=getCheck(getForm("LinkStatus","post"))
	LinkURL=getForm("LinkURL","post")	
	Linkgroup=getForm("Linkgroup","post")	
	
	if isnul(LinkText) then alertMsgAndGo "��վ���Ʋ���Ϊ��","-1"
	if isnul(LinkURL) then alertMsgAndGo "���ӵ�ַ����Ϊ��","-1"		
	if not isurl(LinkURL) then alertMsgAndGo "���ӵ�ַ����ȷ","-1"	
	if isnul(LinkStatus) then LinkStatus=0
	conn.Exec "update {prefix}Links set LinkType="&LinkType&",LinkText='"&LinkText&"',LinkDesc='"&LinkDesc&"',ImageURL='"&ImageURL&"',LinkOrder="&LinkOrder&",LinkStatus="&LinkStatus&",LinkURL='"&LinkURL&"',Linkgroup="&Linkgroup&" where LinkID="&ID,"exe"
	
	alertMsgAndGo "�޸ĳɹ�","AspCms_Links.asp"
End Sub

Sub addLinks 	
	LinkType=getForm("LinkType","post")	
	LinkText=getForm("LinkText","post")	
	LinkDesc=getForm("LinkDesc","post")	
	ImageURL =getForm("ImageURL","post")	
	LinkOrder=getForm("LinkOrder","post")
	LinkStatus=getCheck(getForm("LinkStatus","post"))
	LinkURL=getForm("LinkURL","post")	
	Linkgroup=getForm("Linkgroup","post")	
	
	if isnul(LinkText) then alertMsgAndGo "��վ���Ʋ���Ϊ��","-1"
	if isnul(LinkURL) then alertMsgAndGo "���ӵ�ַ����Ϊ��","-1"		
	if not isurl(LinkURL) then alertMsgAndGo "���ӵ�ַ����ȷ","-1"
	'if LinkType and isnul(ImageURL) then alertMsgAndGo "ͼƬ��ַ����Ϊ��","-1"
	conn.Exec "insert into {prefix}Links(LinkType,LinkText,LinkDesc,ImageURL,LinkOrder,LinkStatus,LinkURL,Linkgroup) values('"&LinkType&"','"&LinkText&"','"&LinkDesc&"','"&ImageURL&"','"&LinkOrder&"','"&LinkStatus&"','"&LinkURL&"',"&Linkgroup&")","exe"		
	alertMsgAndGo "��ӳɹ�","AspCms_Links.asp"
End Sub	

Sub linksList
	Dim rsObj	:	Set rsObj=conn.Exec("select LinkID,LinkText,LinkURL,ImageURL,LinkOrder,LinkStatus,LinkType,Linkgroup from {prefix}Links Order by LinkOrder Asc,LinkID","r1")
	If rsObj.Eof Then 
		echo"<tr bgcolor=""#FFFFFF"" align=""center"">"&vbcrlf& _
			"<td colspan=""7"">û������</td>"&vbcrlf& _
		  "</tr>"&vbcrlf
	Else
		dim i
		i=0
		Do while not rsObj.Eof 
		Dim linkDisplay,linktype
		if rsObj(6) ="0" then linkDisplay=rsObj(1) : linktype ="��������"
		if rsObj(6) ="1" then linkDisplay="<img height=""30"" title="""&rsObj(1)&""" src="""&rsObj(3)&""" />" : linktype ="ͼƬ����"
		 echo"<tr class=list>"&vbcrlf& _
			"<TD align=middle height=26><INPUT type=checkbox value="""&rsObj(0)&""" name=""id""></TD>"&vbcrlf& _
			"<td>"&rsObj(0)&"</td>"&vbcrlf& _
			"<td>"&linkDisplay&"</td>"&vbcrlf& _
			"<td>"&rsObj(2)&"</td>"&vbcrlf& _
			"<td>"&rsObj(4)&"</td>"&vbcrlf& _
			"<td>"&linktype&"</td>"&vbcrlf& _
			"<td><input type=""hidden"" name=""lid"" value="""&rsObj(0)&""" style=""width:40px;"">"
			
        if rsObj(7)=1 then echo "<input name=""Linkgroup"&i&""" type=""radio"" value=""1"" checked=""checked"">1" else echo "<input name=""Linkgroup"&i&""" type=""radio"" value=""1"">1"
      
		 if rsObj(7)=2 then echo "<input name=""Linkgroup"&i&""" type=""radio"" value=""2"" checked=""checked"">2" else echo "<input name=""Linkgroup"&i&""" type=""radio"" value=""2"">2"
		  if rsObj(7)=3 then  echo "<input name=""Linkgroup"&i&""" type=""radio"" value=""3"" checked=""checked"">3" else echo "<input name=""Linkgroup"&i&""" type=""radio"" value=""3"">3"
        echo "</td>"&vbcrlf& _
			"<TD>"&getStr(rsObj(5),"<a href=""?action=off&id="&rsObj(0)&""" title=""����""><IMG src=""../../images/toolbar_ok.gif""></a>","<a href=""?action=on&id="&rsObj(0)&""" title=""����"" ><IMG src=""../../images/toolbar_no.gif""></a>")&"</TD>"&vbcrlf& _
			"<td><a href=""AspCms_LinksEdit.asp?id="&rsObj(0)&""" class=""txt_C1"">�޸�</a> | <a href=""?action=del&id="&rsObj(0)&""" class=""txt_C1"" onClick=""return confirm('ȷ��Ҫɾ����')"">ɾ��</a></td>"&vbcrlf& _
		  "</tr>"&vbcrlf
		  i=i+1
		  rsObj.MoveNext
		Loop
	End If
	rsObj.close	:	Set rsObj = nothing
End Sub

Sub delLinks 
	id=getForm("id","both")
	if isnul(id) then alertMsgAndGo "��ѡ��Ҫɾ��������","-1"

	conn.Exec "delete from {prefix}Links where LinkID in("&id&")" ,"exe"
'	alertMsgAndGo "ɾ���ɹ�","AspCms_Links.asp"	
	response.Redirect("AspCms_Links.asp")
End Sub 

Sub updategroup
	Dim ids				:	ids=split(getForm("lid","post"),",")
	Dim Linkgroups			
	If Ubound(ids)=-1 Then 	'��ֹ��ֵΪ��ʱ�±�Խ��
		ReDim ids(0)
		ids(0)=""
	End If	
	
	Dim i
	
	For i=0 To Ubound(ids)
		Linkgroups=getForm("Linkgroup"&i,"post")		
		if isnum(trim(Linkgroups)) then
			Conn.Exec "update {prefix}links Set LinkGroup="&trim(Linkgroups)&" Where linkID="&trim(ids(i)),"exe"	
		else
			Conn.Exec "update {prefix}links Set LinkGroup=0 Where linkID="&trim(ids(i)),"exe"	
		end if
	Next
	
	
	alertMsgAndGo "������ɹ�",getPageName()
End Sub

%>