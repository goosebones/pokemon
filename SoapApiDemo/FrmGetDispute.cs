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
	/// Summary description for GetDisputeForm.
	/// </summary>
	public class FrmGetDispute : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnGetDispute;
		private System.Windows.Forms.GroupBox GrpResult;
		internal System.Windows.Forms.Label LblTitle;
		private System.Windows.Forms.TextBox TxtTitle;
		private System.Windows.Forms.Label LblDisputeId;
		private System.Windows.Forms.TextBox TxtDisputeId;
		private System.Windows.Forms.TextBox TxtBuyer;
		internal System.Windows.Forms.Label LblBuyer;
		internal System.Windows.Forms.Label LblCreatedTime;
		private System.Windows.Forms.TextBox TxtCreatedTime;
		private System.Windows.Forms.TextBox TxtSeller;
		internal System.Windows.Forms.Label LblSeller;
		internal System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.ListView LstMessages;
		private System.Windows.Forms.ColumnHeader ClmSource;
		private System.Windows.Forms.ColumnHeader ClmTime;
		private System.Windows.Forms.ColumnHeader ClmMessage;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetDispute()
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
			this.LblDisputeId = new System.Windows.Forms.Label();
			this.TxtDisputeId = new System.Windows.Forms.TextBox();
			this.BtnGetDispute = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LstMessages = new System.Windows.Forms.ListView();
			this.ClmSource = new System.Windows.Forms.ColumnHeader();
			this.ClmTime = new System.Windows.Forms.ColumnHeader();
			this.ClmMessage = new System.Windows.Forms.ColumnHeader();
			this.LblItemId = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblTitle = new System.Windows.Forms.Label();
			this.TxtTitle = new System.Windows.Forms.TextBox();
			this.TxtBuyer = new System.Windows.Forms.TextBox();
			this.LblBuyer = new System.Windows.Forms.Label();
			this.LblCreatedTime = new System.Windows.Forms.Label();
			this.TxtCreatedTime = new System.Windows.Forms.TextBox();
			this.TxtSeller = new System.Windows.Forms.TextBox();
			this.LblSeller = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// LblDisputeId
			// 
			this.LblDisputeId.Location = new System.Drawing.Point(112, 8);
			this.LblDisputeId.Name = "LblDisputeId";
			this.LblDisputeId.Size = new System.Drawing.Size(80, 23);
			this.LblDisputeId.TabIndex = 40;
			this.LblDisputeId.Text = "Dispute Id:";
			// 
			// TxtDisputeId
			// 
			this.TxtDisputeId.Location = new System.Drawing.Point(192, 8);
			this.TxtDisputeId.Name = "TxtDisputeId";
			this.TxtDisputeId.Size = new System.Drawing.Size(128, 20);
			this.TxtDisputeId.TabIndex = 27;
			this.TxtDisputeId.Text = "";
			// 
			// BtnGetDispute
			// 
			this.BtnGetDispute.Location = new System.Drawing.Point(192, 40);
			this.BtnGetDispute.Name = "BtnGetDispute";
			this.BtnGetDispute.Size = new System.Drawing.Size(104, 23);
			this.BtnGetDispute.TabIndex = 28;
			this.BtnGetDispute.Text = "GetDispute";
			this.BtnGetDispute.Click += new System.EventHandler(this.BtnGetDispute_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LstMessages,
																					this.LblItemId,
																					this.TxtItemId,
																					this.LblTitle,
																					this.TxtTitle,
																					this.TxtBuyer,
																					this.LblBuyer,
																					this.LblCreatedTime,
																					this.TxtCreatedTime,
																					this.TxtSeller,
																					this.LblSeller});
			this.GrpResult.Location = new System.Drawing.Point(8, 72);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(440, 336);
			this.GrpResult.TabIndex = 41;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LstMessages
			// 
			this.LstMessages.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						  this.ClmSource,
																						  this.ClmTime,
																						  this.ClmMessage});
			this.LstMessages.GridLines = true;
			this.LstMessages.Location = new System.Drawing.Point(8, 168);
			this.LstMessages.Name = "LstMessages";
			this.LstMessages.Size = new System.Drawing.Size(424, 120);
			this.LstMessages.TabIndex = 76;
			this.LstMessages.View = System.Windows.Forms.View.Details;
			// 
			// ClmSource
			// 
			this.ClmSource.Text = "Source";
			this.ClmSource.Width = 83;
			// 
			// ClmTime
			// 
			this.ClmTime.Text = "Time";
			this.ClmTime.Width = 80;
			// 
			// ClmMessage
			// 
			this.ClmMessage.Text = "Message";
			this.ClmMessage.Width = 250;
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(8, 24);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(72, 23);
			this.LblItemId.TabIndex = 74;
			this.LblItemId.Text = "Item Id:";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(80, 24);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.ReadOnly = true;
			this.TxtItemId.Size = new System.Drawing.Size(88, 20);
			this.TxtItemId.TabIndex = 75;
			this.TxtItemId.Text = "";
			// 
			// LblTitle
			// 
			this.LblTitle.Location = new System.Drawing.Point(8, 48);
			this.LblTitle.Name = "LblTitle";
			this.LblTitle.Size = new System.Drawing.Size(72, 23);
			this.LblTitle.TabIndex = 61;
			this.LblTitle.Text = "Title:";
			// 
			// TxtTitle
			// 
			this.TxtTitle.Location = new System.Drawing.Point(80, 48);
			this.TxtTitle.Name = "TxtTitle";
			this.TxtTitle.ReadOnly = true;
			this.TxtTitle.Size = new System.Drawing.Size(344, 20);
			this.TxtTitle.TabIndex = 72;
			this.TxtTitle.Text = "";
			// 
			// TxtBuyer
			// 
			this.TxtBuyer.Location = new System.Drawing.Point(80, 96);
			this.TxtBuyer.Name = "TxtBuyer";
			this.TxtBuyer.ReadOnly = true;
			this.TxtBuyer.Size = new System.Drawing.Size(120, 20);
			this.TxtBuyer.TabIndex = 73;
			this.TxtBuyer.Text = "";
			// 
			// LblBuyer
			// 
			this.LblBuyer.Location = new System.Drawing.Point(8, 96);
			this.LblBuyer.Name = "LblBuyer";
			this.LblBuyer.Size = new System.Drawing.Size(72, 23);
			this.LblBuyer.TabIndex = 60;
			this.LblBuyer.Text = "Buyer:";
			// 
			// LblCreatedTime
			// 
			this.LblCreatedTime.Location = new System.Drawing.Point(8, 120);
			this.LblCreatedTime.Name = "LblCreatedTime";
			this.LblCreatedTime.Size = new System.Drawing.Size(72, 23);
			this.LblCreatedTime.TabIndex = 62;
			this.LblCreatedTime.Text = "Created:";
			// 
			// TxtCreatedTime
			// 
			this.TxtCreatedTime.Location = new System.Drawing.Point(80, 120);
			this.TxtCreatedTime.Name = "TxtCreatedTime";
			this.TxtCreatedTime.ReadOnly = true;
			this.TxtCreatedTime.Size = new System.Drawing.Size(120, 20);
			this.TxtCreatedTime.TabIndex = 69;
			this.TxtCreatedTime.Text = "";
			// 
			// TxtSeller
			// 
			this.TxtSeller.Location = new System.Drawing.Point(80, 72);
			this.TxtSeller.Name = "TxtSeller";
			this.TxtSeller.ReadOnly = true;
			this.TxtSeller.Size = new System.Drawing.Size(120, 20);
			this.TxtSeller.TabIndex = 70;
			this.TxtSeller.Text = "";
			// 
			// LblSeller
			// 
			this.LblSeller.Location = new System.Drawing.Point(8, 72);
			this.LblSeller.Name = "LblSeller";
			this.LblSeller.Size = new System.Drawing.Size(72, 23);
			this.LblSeller.TabIndex = 54;
			this.LblSeller.Text = "Seller:";
			// 
			// FrmGetDispute
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(456, 445);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.GrpResult,
																		  this.LblDisputeId,
																		  this.TxtDisputeId,
																		  this.BtnGetDispute});
			this.Name = "FrmGetDispute";
			this.Text = "GetDispute";
			this.Load += new System.EventHandler(this.FrmGetDispute_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetDispute_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstMessages.Items.Clear();
				TxtItemId.Text =  "";
				TxtTitle.Text =  "";
				TxtSeller.Text =  "";
				TxtBuyer.Text =  "";
				TxtCreatedTime.Text =  "";

				GetDisputeCall apicall = new GetDisputeCall(Context);
				DisputeType dispute = apicall.GetDispute(TxtDisputeId.Text);
			
				TxtItemId.Text = dispute.Item.ItemID;
				TxtTitle.Text = dispute.Item.Title;
				TxtSeller.Text = dispute.SellerUserID;
				TxtBuyer.Text = dispute.BuyerUserID;
				TxtCreatedTime.Text = dispute.DisputeCreatedTime.ToString();

				foreach (DisputeMessageType message in dispute.DisputeMessage)
				{
					string[] listparams = new string[5];
					listparams[0] = message.MessageSource.ToString();
					listparams[1] = message.MessageCreationTime.ToString();
					listparams[2] = message.MessageText;

					ListViewItem vi = new ListViewItem(listparams);
					this.LstMessages.Items.Add(vi);

				}

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void FrmGetDispute_Load(object sender, System.EventArgs e)
		{
		
		}

	
	}
}
