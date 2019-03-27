<%@ Page language="c#" Inherits="iWebOffice2000.ocx.c.net.DocumentList" CodeFile="DocumentList.aspx.cs" %>
<HTML>
<title>���Ƽ�iWebOffice2000ʵ������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="test.css" type="text/css" rel="stylesheet">

<script language="javascript">
//���ã���ʾ����״̬
function StatusMsg(mString){
  StatusBar.innerHTML=mString;
}

//��ʼ������
function initObject(){
  WebOffice.WebUrl="OfficeServer.aspx";      //��̨����ҳ·��������ִ�к�̨���ݴ���ҵ�񡣱�����֧�����·��
  WebOffice.RecordID="1234";                //�ĵ���¼��
  WebOffice.Template="";                    //ģ���¼��
  WebOffice.FileName="1234.doc";            //�ĵ�����
  WebOffice.FileType=".doc";                //�ĵ�����  .doc  .xls
  WebOffice.UserName="�ĵ��༭��";          //��ǰ����Ա
  WebOffice.EditType="1";                   //�༭״̬����һλ����Ϊ0,1,2,3���У�0���ɱ༭��1���Ա༭,�޺ۼ���2���Ա༭,�кۼ�,�����޶���3���Ա༭,�кۼ�,���޶���

  WebOffice.MaxFileSize = 4 * 1024;         //�����ĵ���С���ƣ�Ĭ����8M���������ó�4M��
  WebOffice.ShowMenu="1";                   //�Ƿ���ʾ�˵���1��ʾ��0����ʾ

  WebOffice.CreateFile();                   //�����հ��ĵ�
}

//���ã��򿪷������ĵ�
function LoadDocument(){
  if (!WebOffice.WebOpen()){                //�򿪸��ĵ�    ����OfficeServer��OPTION="LOADFILE"
    StatusMsg(WebOffice.Status);            //��ʾ״̬����OfficeServer�ж�ȡ
  }else{
    StatusMsg(WebOffice.Status);            //��ʾ״̬����OfficeServer�ж�ȡ
  }
}

//���ã�����������ĵ�
function SaveDocument(){
  if (!WebOffice.WebSave()){     //����OfficeServer��OPTION="SAVEFILE"
    StatusMsg(WebOffice.Status);
  }else{
    StatusMsg(WebOffice.Status);
  }
}

//���ã������հ��ĵ�
function CreateFile(){
  WebOffice.CreateFile();
  StatusMsg(WebOffice.Status);
}

//���ã��򿪱����ļ�
function WebOpenLocal(){
  var result = WebOffice.WebOpenLocalFile();
  if(result){
    StatusMsg("�򿪱����ĵ��ɹ���");
  }else{
    StatusMsg(WebOffice.Status);
  }
}

//���ã������ļ�������
function WebSaveLocal(){
  var result = WebOffice.WebSaveLocalFile();
  if(result){
    StatusMsg("�����ĵ������سɹ���");
  }else{
    StatusMsg(WebOffice.Status);
  }
}

//���ã���ȡ�ĵ�ҳ����VBA��չӦ�ã�
function WebDocumentPageCount(){
  if (WebOffice.FileType==".doc"){
    var intPageTotal = WebOffice.WebObject.Application.ActiveDocument.BuiltInDocumentProperties(14);
    alert("�ĵ�ҳ������"+intPageTotal);
  }
  
  if (WebOffice.FileType==".wps"){
    var intPageTotal = WebOffice.WebObject.PagesCount();
    alert("�ĵ�ҳ������"+intPageTotal);
  }
}

//���ã������ĵ���ȫ���ۼ���VBA��չӦ�ã�
function WebAcceptAllRevisions(){
  WebOffice.WebObject.Application.ActiveDocument.AcceptAllRevisions();
  var mCount = WebOffice.WebObject.Application.ActiveDocument.Revisions.Count;
  if(mCount>0){
    StatusMsg("���ܺۼ�ʧ�ܣ�");
    return false;
  }else{
    StatusMsg("�ĵ��еĺۼ��Ѿ�ȫ�����ܣ�");
    return true;
  }
}

//���ã��˳�iWebOffice
function UnLoad(){
  try{
    if (!WebOffice.WebClose()){
      StatusMsg(WebOffice.Status);
    }
    else{
      StatusMsg("�ر��ĵ�...");
    }
  }
  catch(e){
    alert(e.description);
  }
}


</script>

<BODY onLoad="initObject();" onunload="UnLoad()"><!--�������ļ�-->
  <div style="text-align:center;font-size:18px">iWebOffice2000�����ĵ��м��[��Ѱ�]ʵ������</div>
  <hr size=1>  
  <div style="font-size:12px;color:#FF0000;">
    ��ע��iWebOffice2000�ؼ������ߴ򿪡��༭�ͱ���������ϵ��ĵ�����Ҫ��ʾ��ʾ������Ҫ��<br>
    ������1.�밲װOffice2000���ϰ汾�������ʹ�ô�IE�����������ҳ��ʱ��װiWebOffice2000�����<br>
    ������2.��������Զ���װiWebOffice2000�����������������<a href="InstallClient.zip">[���ذ�װ����]</a>����װ��
  </div>
  <hr size=1>
  <input type=button class=button value="�½��ļ�" onclick="CreateFile()">
  <input type=button class=button value="�༭����" onClick="WebOffice.EditType='2'">
  <input type=button class=button value="�༭������" onClick="WebOffice.EditType='1'">
��<input type=button class=button value="�򿪱����ļ�" onclick="WebOpenLocal()">
  <input type=button class=button value="���汾���ļ�" onclick="WebSaveLocal()">
  <input type=button class=button value="��Զ���ļ�" onClick="LoadDocument()"><!--ϵͳ��ͨ��WebUrlָ���ĳ��򵽷������ϵ����ļ�, ������RecordIDָ�����ļ�-->
  <input type=button class=button value="����Զ���ļ�" onClick="SaveDocument()"><!--ϵͳ��ͨ��WebUrlָ���ĳ��򱣴汾�ļ�����������,������RecordIDָ�����ļ�-->
��<input type=button class=button value="�ĵ�ҳ��" onClick="WebDocumentPageCount()">
  <input type=button class=button value="���ܺۼ�" onClick="WebAcceptAllRevisions()">
  <table>
    <tr>
      <td height="28" style="font-size:12px;color:#0000FF">״̬����</td>
      <td id=StatusBar style="font-size:12px;color:#FF0000">����״̬��Ϣ</td>
    </tr>
  </table>
  <!--����iWebOffice��ע��汾�ţ�����������-->
  <script src="iWebOffice2000.js"></script>
</BODY>
</HTML>