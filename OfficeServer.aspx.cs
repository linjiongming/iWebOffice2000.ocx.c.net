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
	/// OfficeServer 的摘要说明。
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
			// 在此处放置用户代码以初始化页面

			MsgObj = new DBstep.iMsgServer2000();
			mFilePath=Server.MapPath(".");

			MsgObj.MsgVariant(Request.BinaryRead(Request.ContentLength));
			if( MsgObj.GetMsgByName("DBSTEP").Equals("DBSTEP"))                                //如果是合法的信息包
			{
				mOption=MsgObj.GetMsgByName("OPTION");                                         //取得操作信息
				mUserName=MsgObj.GetMsgByName("USERNAME");									   //取得操作用户名称
				if(mOption.Equals("LOADFILE"))						                           //下面的代码为打开服务器数据库里的文件
				{
					mRecordID=MsgObj.GetMsgByName("RECORDID");		                          //取得文档编号
					mFileType=MsgObj.GetMsgByName("FILETYPE");	                              //取得文档类型
					MsgObj.MsgTextClear();                                                    //清除文本信息
					mFilePath=mFilePath+"/Document/"+mRecordID+mFileType;                     //全文批注文件的完整路径
					if (MsgObj.MsgFileLoad(mFilePath))				                          //调入文档
					{
						//MsgObj.MsgFileBody(mFileBody);					                  //将文件信息打包，mFileBody为从数据库中读取，类型byte[]
						MsgObj.SetMsgByName("STATUS","打开全文批注成功!");		              //设置状态信息
						MsgObj.MsgError("");		                                          //清除错误信息
					}
					else
					{
						MsgObj.MsgError("打开全文批注失败!");		                          //设置错误信息
					}
				}
				else if(mOption.Equals("SAVEFILE"))					                          //下面的代码为保存文件在服务器的数据库里
				{
					mRecordID=MsgObj.GetMsgByName("RECORDID");		                          //取得文档编号
					mFileType=MsgObj.GetMsgByName("FILETYPE");		                          //取得文档类型
					//mMyDefine1=MsgObj.GetMsgByName("MyDefine1");	                          //取得客户端传递变量值 MyDefine1="自定义变量值1"
					//mFileBody=MsgObj.MsgFileBody();	                                      //取得文档内容 mFileBody可以保存到数据库中，类型byte[]
					MsgObj.MsgTextClear();                                                    //清除文本信息
					mFilePath=mFilePath+"/Document/"+mRecordID+mFileType;                     //全文批注文件的完整路径
					if (MsgObj.MsgFileSave(mFilePath))				                          //保存文档内容
					{
						MsgObj.SetMsgByName("STATUS", "保存全文批注成功!");	                  //设置状态信息
						MsgObj.MsgError("");						                          //清除错误信息
					}
					else
					{
						MsgObj.MsgError("保存全文批注失败!");		                           //设置错误信息
					}
					MsgObj.MsgFileClear();
				}

			}
			else
			{				
				MsgObj.MsgError("客户端发送数据包错误!");
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
			// CODEGEN：该调用是 ASP.NET Web 窗体设计器所必需的。
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// 设计器支持所需的方法 - 不要使用代码编辑器修改
		/// 此方法的内容。
		/// </summary>
		private void InitializeComponent()
		{    
		}
		#endregion
	}
}
