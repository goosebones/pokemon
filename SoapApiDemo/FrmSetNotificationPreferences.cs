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
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using eBay.Service.Core.Soap;
using eBay.Service.Core.Sdk;
using eBay.Service.Call;
using eBay.Service.Util;


namespace SoapLibraryDemo
{
	/// <summary>
	/// Summary description for AddItemForm.
	/// </summary>
	public class FrmSetNotificationPreferences : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnSetNotificationPreferences;
		private System.Windows.Forms.ComboBox CboStatus;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.TextBox TxtUrl;
		private System.Windows.Forms.Label LblUrl;
		private System.Windows.Forms.Label[] LblEvents;
		private System.Windows.Forms.ComboBox[] CboEventStatus;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.Panel PnlUserPrefs;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmSetNotificationPreferences()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
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
			this.BtnSetNotificationPreferences = new System.Windows.Forms.Button();
			this.CboStatus = new System.Windows.Forms.ComboBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtUrl = new System.Windows.Forms.TextBox();
			this.LblUrl = new System.Windows.Forms.Label();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.PnlUserPrefs = new System.Windows.Forms.Panel();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnSetNotificationPreferences
			// 
			this.BtnSetNotificationPreferences.Location = new System.Drawing.Point(168, 344);
			this.BtnSetNotificationPreferences.Name = "BtnSetNotificationPreferences";
			this.BtnSetNotificationPreferences.Size = new System.Drawing.Size(184, 26);
			this.BtnSetNotificationPreferences.TabIndex = 46;
			this.BtnSetNotificationPreferences.Text = "SetNotificationPreferences";
			this.BtnSetNotificationPreferences.Click += new System.EventHandler(this.BtnSetNotificationPreferences_Click);
			// 
			// CboStatus
			// 
			this.CboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboStatus.Location = new System.Drawing.Point(232, 32);
			this.CboStatus.Name = "CboStatus";
			this.CboStatus.Size = new System.Drawing.Size(144, 21);
			this.CboStatus.TabIndex = 59;
			// 
			// LblStatus
			// 
			this.LblStatus.Location = new System.Drawing.Point(16, 32);
			this.LblStatus.Name = "LblStatus";
			this.LblStatus.Size = new System.Drawing.Size(216, 18);
			this.LblStatus.TabIndex = 58;
			this.LblStatus.Text = "Application Status:";
			// 
			// TxtUrl
			// 
			this.TxtUrl.Location = new System.Drawing.Point(232, 8);
			this.TxtUrl.Name = "TxtUrl";
			this.TxtUrl.Size = new System.Drawing.Size(256, 20);
			this.TxtUrl.TabIndex = 56;
			this.TxtUrl.Text = "http://www.ebay.com";
			// 
			// LblUrl
			// 
			this.LblUrl.Location = new System.Drawing.Point(16, 8);
			this.LblUrl.Name = "LblUrl";
			this.LblUrl.Size = new System.Drawing.Size(216, 16);
			this.LblUrl.TabIndex = 57;
			this.LblUrl.Text = "Delivery URL:";
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.label1);
			this.GrpResult.Controls.Add(this.TxtStatus);
			this.GrpResult.Location = new System.Drawing.Point(8, 376);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(488, 64);
			this.GrpResult.TabIndex = 60;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 23);
			this.label1.TabIndex = 42;
			this.label1.Text = "Status:";
			// 
			// TxtStatus
			// 
			this.TxtStatus.Location = new System.Drawing.Point(96, 24);
			this.TxtStatus.Name = "TxtStatus";
			this.TxtStatus.ReadOnly = true;
			this.TxtStatus.Size = new System.Drawing.Size(72, 20);
			this.TxtStatus.TabIndex = 41;
			this.TxtStatus.Text = "";
			// 
			// PnlUserPrefs
			// 
			this.PnlUserPrefs.AutoScroll = true;
			this.PnlUserPrefs.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.PnlUserPrefs.Location = new System.Drawing.Point(8, 64);
			this.PnlUserPrefs.Name = "PnlUserPrefs";
			this.PnlUserPrefs.Size = new System.Drawing.Size(488, 272);
			this.PnlUserPrefs.TabIndex = 61;
			// 
			// FrmSetNotificationPreferences
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 453);
			this.Controls.Add(this.PnlUserPrefs);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.CboStatus);
			this.Controls.Add(this.LblStatus);
			this.Controls.Add(this.TxtUrl);
			this.Controls.Add(this.LblUrl);
			this.Controls.Add(this.BtnSetNotificationPreferences);
			this.Name = "FrmSetNotificationPreferences";
			this.Text = "SetNotificationPreferences";
			this.Load += new System.EventHandler(this.FrmSetNotificationPreferences_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
	
	
		private void FrmSetNotificationPreferences_Load(object sender, System.EventArgs e)
		{
			string[] items = Enum.GetNames(typeof(EnableCodeType));
			string[] stats = new string[items.Length -1];

			int index = 0;
			foreach (string stat in items)
			{
				if (stat != "CustomCode")
				{
					stats[index] = stat;
					index++;
				}
			}

	
			CboStatus.Items.AddRange(stats);;
			CboStatus.SelectedIndex = 0;


			//string[] events = Enum.GetNames(typeof(NotificationEventTypeCodeType));
			string[] events = {
				NotificationEventTypeCodeType.OutBid.ToString(),
				NotificationEventTypeCodeType.EndOfAuction.ToString(),
				NotificationEventTypeCodeType.AuctionCheckoutComplete.ToString(),
				NotificationEventTypeCodeType.CheckoutBuyerRequestsTotal.ToString(),
			    NotificationEventTypeCodeType.Feedback.ToString(),
				NotificationEventTypeCodeType.FeedbackForSeller.ToString(),
				NotificationEventTypeCodeType.SecondChanceOffer.ToString(),
				NotificationEventTypeCodeType.AskSellerQuestion.ToString(),
				NotificationEventTypeCodeType.ItemListed.ToString(),
				NotificationEventTypeCodeType.ItemRevised.ToString(),
				NotificationEventTypeCodeType.BuyerResponseDispute.ToString(),
				NotificationEventTypeCodeType.SellerRespondedToDispute.ToString(),
				NotificationEventTypeCodeType.SellerClosedDispute.ToString(),
				NotificationEventTypeCodeType.BestOffer.ToString(),
				NotificationEventTypeCodeType.MyMessageseBayMessage.ToString()
			};

			LblEvents = new System.Windows.Forms.Label[events.Length];
			CboEventStatus = new System.Windows.Forms.ComboBox[events.Length];
			
			int y = 8;
			int x = 0;
			foreach (string ev in events)
			{
				if (ev != "CustomCode" && ev != "None")
				{
					System.Windows.Forms.Label lbl = new System.Windows.Forms.Label();
					System.Windows.Forms.ComboBox cbo = new System.Windows.Forms.ComboBox();
					lbl.Location = new System.Drawing.Point(8, y);
					lbl.Size = new System.Drawing.Size(160, 16);
					lbl.Text = ev.ToString();

					cbo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
					cbo.Location = new System.Drawing.Point(232, y);
					cbo.Size = new System.Drawing.Size(144, 21);
					cbo.Items.AddRange(stats);
					cbo.SelectedIndex = 0;

					LblEvents[x] = lbl;
					CboEventStatus[x] = cbo;
					
					y = y + 24;
					x++;
				}
				PnlUserPrefs.Controls.AddRange(LblEvents);
				PnlUserPrefs.Controls.AddRange(CboEventStatus);
			}

		}

		private void BtnSetNotificationPreferences_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";
	
				SetNotificationPreferencesCall apicall = new SetNotificationPreferencesCall(Context);
				apicall.ApplicationDeliveryPreferences = new ApplicationDeliveryPreferencesType();
				apicall.ApplicationDeliveryPreferences.ApplicationEnable = (EnableCodeType) Enum.Parse(typeof(EnableCodeType), CboStatus.SelectedItem.ToString());
				
				if (TxtUrl.Text.Length > 0)
					apicall.ApplicationDeliveryPreferences.ApplicationURL = TxtUrl.Text;

				NotificationEnableTypeCollection notifications = new NotificationEnableTypeCollection();
				for (int inx = 0; inx < LblEvents.Length; inx++)
				{
					NotificationEnableType net = new NotificationEnableType();
					net.EventType = (NotificationEventTypeCodeType) Enum.Parse(typeof(NotificationEventTypeCodeType), LblEvents[inx].Text);
					net.EventEnable = (EnableCodeType) Enum.Parse(typeof(EnableCodeType), CboEventStatus[inx].SelectedItem.ToString());
					notifications.Add(net);
				}

				apicall.SetNotificationPreferences(notifications);
	
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();

			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}
}
