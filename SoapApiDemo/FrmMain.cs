#region Copyright
//	Copyright (c) 2013 eBay, Inc.
//	
//	This program is licensed under the terms of the eBay Common Development and
//	Distribution License (CDDL) Version 1.0 (the "License") and any subsequent  
//	version thereof released by eBay.  The then-current version of the License can be 
//	found at http://www.opensource.org/licenses/cddl1.php and in the eBaySDKLicense 
//	file that is under the eBay SDK ../docs directory
#endregion

using System;
using System.Net;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using eBay.Service.Core.Sdk;
using Samples.Helper;

namespace SoapLibraryDemo
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class FrmMain : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnAccountSettings;
		private System.Windows.Forms.ListBox APICallListBox;
		private System.Windows.Forms.Button runbutton;
		
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmMain()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();			
			AddCommandToListBox("AddItem", typeof(FrmAddItem));
			AddCommandToListBox("AddSecondChanceItem", typeof(FrmAddSecondChanceItem));			
			AddCommandToListBox("GetAccount", typeof(FrmGetAccount));
			AddCommandToListBox("GetSellerEvents", typeof(FrmGetSellerEvents));
			AddCommandToListBox("GetUser", typeof(FrmGetUser));
			AddCommandToListBox("LeaveFeedback", typeof(FrmLeaveFeedback));

			AddCommandToListBox("GetTokenStatus", typeof(FrmGetTokenStatus));

			AddCommandToListBox("AddToItemDescription", typeof(FrmAddToItemDescription));
			AddCommandToListBox("GetItemTransactions", typeof(FrmGetItemTransactions));
			AddCommandToListBox("GetItem", typeof(FrmGetItem));
			AddCommandToListBox("GetCategories", typeof(FrmGetCategories));
			AddCommandToListBox("GetFeedback", typeof(FrmGetFeedback));
			AddCommandToListBox("GetMyeBaySelling", typeof(FrmGetMyeBaySelling));
			AddCommandToListBox("GetMyeBayBuying", typeof(FrmGetMyeBayBuying));

			AddCommandToListBox("EndItem", typeof(FrmEndItem));


			AddCommandToListBox("RelistItem", typeof(FrmRelistItem));
			AddCommandToListBox("GetSellerTransactions", typeof(FrmGetSellerTransactions));
			AddCommandToListBox("GetSellerList", typeof(FrmGetSellerList));			
			AddCommandToListBox("ReviseItem", typeof(FrmReviseItem));			
			AddCommandToListBox("GetMemberMessages", typeof(FrmGetMemberMessages));


			AddCommandToListBox("AddDispute", typeof(FrmAddDispute));
			AddCommandToListBox("AddDisputeResponse", typeof(FrmAddDisputeResponse));
			AddCommandToListBox("SellerReverseDispute", typeof(FrmSellerReverseDispute));
			AddCommandToListBox("AddOrder", typeof(FrmAddOrder));
			AddCommandToListBox("GetAllBidders", typeof(FrmGetAllBidders));
			AddCommandToListBox("GetApiAccessRules", typeof(FrmGetApiAccessRules));


			AddCommandToListBox("GetPromotionRules", typeof(FrmGetPromotionRules));			
			AddCommandToListBox("GetStore", typeof(FrmGetStore));
			AddCommandToListBox("GetOrders", typeof(FrmGetOrders));
			AddCommandToListBox("SetNotificationPreferences", typeof(FrmSetNotificationPreferences));
			AddCommandToListBox("GetCrossPromotions", typeof(FrmGetCrossPromotions));
			AddCommandToListBox("GetDispute", typeof(FrmGetDispute));


			AddCommandToListBox("GeteBayOfficialTime", typeof(FrmGeteBayOfficialTime));
			AddCommandToListBox("GetItemShipping", typeof(FrmGetItemShipping));
			AddCommandToListBox("GetNotificationPreferences", typeof(FrmGetNotificationPreferences));
			AddCommandToListBox("GetStoreOptions", typeof(FrmGetStoreOptions));
			AddCommandToListBox("GetSuggestedCategories", typeof(FrmGetSuggestedCategories));
			AddCommandToListBox("ReviseCheckoutStatus", typeof(FrmReviseCheckoutStatus));
			AddCommandToListBox("GetUserDisputes", typeof(FrmGetUserDisputes));
			AddCommandToListBox("GetStoreCustomPage", typeof(FrmGetStoreCustomPage));
			//AddCommandToListBox("GetHighBidders", typeof(FrmGetHighBidders));

			AddCommandToListBox("GetBestOffers", typeof(FrmGetBestOffers));
			AddCommandToListBox("RespondToBestOffer", typeof(FrmRespondToBestOffer));
			AddCommandToListBox("SentInvoice", typeof(FrmSendInvoice));
			AddCommandToListBox("SetStore", typeof(FrmSetStore));
			
			//AddCommandToListBox("GetItemRecommendations", typeof(FrmGetItemRecommendations));
            AddCommandToListBox("UploadPictures", typeof(FrmUploadPictures));
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			SetFormStartPosition();
		}

		private void SetFormStartPosition()
		{
			this.StartPosition = FormStartPosition.CenterScreen;
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.BtnAccountSettings = new System.Windows.Forms.Button();
			this.APICallListBox = new System.Windows.Forms.ListBox();
			this.runbutton = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// BtnAccountSettings
			// 
			this.BtnAccountSettings.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.BtnAccountSettings.Location = new System.Drawing.Point(112, 16);
			this.BtnAccountSettings.Name = "BtnAccountSettings";
			this.BtnAccountSettings.Size = new System.Drawing.Size(192, 24);
			this.BtnAccountSettings.TabIndex = 0;
			this.BtnAccountSettings.Text = "Config API Account";
			this.BtnAccountSettings.Click += new System.EventHandler(this.BtnAccountSettings_Click);
			// 
			// APICallListBox
			// 
			this.APICallListBox.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.APICallListBox.Location = new System.Drawing.Point(48, 60);
			this.APICallListBox.Name = "APICallListBox";
			this.APICallListBox.Size = new System.Drawing.Size(336, 186);
			this.APICallListBox.Sorted = true;
			this.APICallListBox.TabIndex = 1;
			this.APICallListBox.DoubleClick += new System.EventHandler(this.APICallListBox_DoubleClick);
			this.APICallListBox.SelectedIndexChanged += new System.EventHandler(this.APICallListBox_SelectedIndexChanged);
			// 
			// runbutton
			// 
			this.runbutton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.runbutton.Location = new System.Drawing.Point(120, 268);
			this.runbutton.Name = "runbutton";
			this.runbutton.Size = new System.Drawing.Size(192, 24);
			this.runbutton.TabIndex = 2;
			this.runbutton.Text = "Run an eBay API call";
			this.runbutton.Click += new System.EventHandler(this.runbutton_Click);
			// 
			// FrmMain
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(424, 309);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.runbutton,
																		  this.APICallListBox,
																		  this.BtnAccountSettings});
			this.ForeColor = System.Drawing.Color.Black;
			this.Name = "FrmMain";
			this.Text = "eBay Soap SDK Sample - API Library Demo (C#)";
			this.Load += new System.EventHandler(this.FrmMain_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new FrmMain());
		}

		private void FrmMain_Load(object sender, System.EventArgs e)
		{			
			Context = AppSettingHelper.GetApiContext();
			Context.ApiLogManager = new ApiLogManager();
			LoggingProperties logProps = AppSettingHelper.GetLoggingProperties();
			Context.ApiLogManager.ApiLoggerList.Add(new eBay.Service.Util.FileLogger(logProps.LogFileName, true, true, true));
			Context.ApiLogManager.EnableLogging = true;
			Context.ApiLogManager.MessageLoggingFilter = getExceptionFilter(logProps);
			Context.Site = eBay.Service.Core.Soap.SiteCodeType.US;
            //set proxy server if necessary
            SetProxy();
		}

        private void SetProxy()
        {
            IWebProxy proxy = null;
            string proxyHost = System.Configuration.ConfigurationManager.AppSettings.Get("Proxy.Host");
            string proxyPort = System.Configuration.ConfigurationManager.AppSettings.Get("Proxy.Port");
            if (proxyHost.Length > 0 && proxyPort.Length > 0)
            {
                proxy = new WebProxy(proxyHost, int.Parse(proxyPort));
                
                string username = System.Configuration.ConfigurationManager.AppSettings.Get("Proxy.Username");
                string password = System.Configuration.ConfigurationManager.AppSettings.Get("Proxy.Password");
                if (username.Length > 0 && password.Length > 0)
                {
                    proxy.Credentials = new NetworkCredential(username, password);
                }
            }

            Context.WebProxy = proxy;
        }

		private ExceptionFilter getExceptionFilter(LoggingProperties logProps)
		{
			if (logProps.LogPayloadErrorCodes == null && logProps.LogPayloadExceptions == null && logProps.LogPayloadHttpStatusCodes == null)
				return null;
			else
				return new ExceptionFilter(logProps.LogPayloadErrorCodes, logProps.LogPayloadExceptions, logProps.LogPayloadHttpStatusCodes);
			
		}

		private void BtnAccountSettings_Click(object sender, System.EventArgs e)
		{
			FrmAccount form = new FrmAccount();
			form.Context = Context;
			if (form.ShowDialog() == DialogResult.OK)
				Context = form.Context;

		}

		private void runbutton_Click(object sender, System.EventArgs e)
		{
			RunSelectedCommand();
		}		
		
		private void RunSelectedCommand()
		{
			CommandListBoxEntry commandListBoxEntry = 
				(CommandListBoxEntry) this.APICallListBox.SelectedItem;
			if(commandListBoxEntry != null)
			{
				System.Type formType = commandListBoxEntry.FormType;
				System.Reflection.ConstructorInfo ci =
					formType.GetConstructor(System.Type.EmptyTypes);
				using(Form form = (Form) ci.Invoke(null))
				{
					form.GetType().GetField("Context").SetValue(form, Context);
					form.ShowDialog();
				}
			}
		}

		private void AddCommandToListBox(string commandName,System.Type formType)
		{
			this.APICallListBox.Items.Add(new CommandListBoxEntry(commandName, formType));
		}


		public class CommandListBoxEntry
		{
			public CommandListBoxEntry(string name, System.Type formType)
			{
				this.Name = name;
				this.FormType = formType;
			}

			public string Name = string.Empty;
			public System.Type FormType = null;

			public override string ToString()
			{
				return this.Name;
			}

		}		
		
		private void APICallListBox_DoubleClick(object sender, System.EventArgs e)
		{
			RunSelectedCommand();
		}


		private void APICallListBox_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			// TODO: Change the text of the Run button.
			this.runbutton.Text = "Run " + this.APICallListBox.SelectedItem.ToString();
		}
	}
}
