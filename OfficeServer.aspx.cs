using System;
using System.IO;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;


namespace iWebOffice2000.ocx.c.net
{
	/// <summary>
	/// OfficeServer ��ժҪ˵����
	/// </summary>
	public partial class OfficeServer : System.Web.UI.Page
	{
		private string mFileType;
		private string mRecordID;
		private string mOption;
		private string mFilePath;

		private string mUserName;
		private DBstep.iMsgServer2000 MsgObj;

		protected void Page_Load(object sender, System.EventArgs e)
		{
			// �ڴ˴������û������Գ�ʼ��ҳ��

			MsgObj = new DBstep.iMsgServer2000();
			mFilePath=Server.MapPath(".");

			MsgObj.MsgVariant(Request.BinaryRead(Request.ContentLength));
			if( MsgObj.GetMsgByName("DBSTEP").Equals("DBSTEP"))                                //����ǺϷ�����Ϣ��
			{
				mOption=MsgObj.GetMsgByName("OPTION");                                         //ȡ�ò�����Ϣ
				mUserName=MsgObj.GetMsgByName("USERNAME");									   //ȡ�ò����û�����
				if(mOption.Equals("LOADFILE"))						                           //����Ĵ���Ϊ�򿪷��������ݿ�����ļ�
				{
					mRecordID=MsgObj.GetMsgByName("RECORDID");		                          //ȡ���ĵ����
					mFileType=MsgObj.GetMsgByName("FILETYPE");	                              //ȡ���ĵ�����
					MsgObj.MsgTextClear();                                                    //����ı���Ϣ
					mFilePath=mFilePath+"/Document/"+mRecordID+mFileType;                     //ȫ����ע�ļ�������·��
					if (MsgObj.MsgFileLoad(mFilePath))				                          //�����ĵ�
					{
						//MsgObj.MsgFileBody(mFileBody);					                  //���ļ���Ϣ�����mFileBodyΪ�����ݿ��ж�ȡ������byte[]
						MsgObj.SetMsgByName("STATUS","��ȫ����ע�ɹ�!");		              //����״̬��Ϣ
						MsgObj.MsgError("");		                                          //���������Ϣ
					}
					else
					{
						MsgObj.MsgError("��ȫ����עʧ��!");		                          //���ô�����Ϣ
					}
				}
				else if(mOption.Equals("SAVEFILE"))					                          //����Ĵ���Ϊ�����ļ��ڷ����������ݿ���
				{
					mRecordID=MsgObj.GetMsgByName("RECORDID");		                          //ȡ���ĵ����
					mFileType=MsgObj.GetMsgByName("FILETYPE");		                          //ȡ���ĵ�����
					//mMyDefine1=MsgObj.GetMsgByName("MyDefine1");	                          //ȡ�ÿͻ��˴��ݱ���ֵ MyDefine1="�Զ������ֵ1"
					//mFileBody=MsgObj.MsgFileBody();	                                      //ȡ���ĵ����� mFileBody���Ա��浽���ݿ��У�����byte[]
					MsgObj.MsgTextClear();                                                    //����ı���Ϣ
					mFilePath=mFilePath+"/Document/"+mRecordID+mFileType;                     //ȫ����ע�ļ�������·��
					if (MsgObj.MsgFileSave(mFilePath))				                          //�����ĵ�����
					{
						MsgObj.SetMsgByName("STATUS", "����ȫ����ע�ɹ�!");	                  //����״̬��Ϣ
						MsgObj.MsgError("");						                          //���������Ϣ
					}
					else
					{
						MsgObj.MsgError("����ȫ����עʧ��!");		                           //���ô�����Ϣ
					}
					MsgObj.MsgFileClear();
				}

			}
			else
			{				
				MsgObj.MsgError("�ͻ��˷������ݰ�����!");
				MsgObj.MsgTextClear();
				MsgObj.MsgFileClear();
			}
            byte[] arr = MsgObj.MsgVariant();
            
            MsgObj = null;

            Response.BinaryWrite(arr);
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN���õ����� ASP.NET Web ���������������ġ�
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// �����֧������ķ��� - ��Ҫʹ�ô���༭���޸�
		/// �˷��������ݡ�
		/// </summary>
		private void InitializeComponent()
		{    
		}
		#endregion
	}
}
