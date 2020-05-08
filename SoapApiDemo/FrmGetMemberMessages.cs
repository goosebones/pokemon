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
	/// Summary description for FormGetMemberMessages.
	/// </summary>
	public class FrmGetMemberMessages : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.DateTimePicker DatePickStartTo;
		private System.Windows.Forms.Label LblStartSep;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.DateTimePicker DatePickStartFrom;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.Button BtnGetMemberMessages;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.ComboBox CboStatus;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.ComboBox CboType;
		private System.Windows.Forms.Label LblType;
		private System.Windows.Forms.CheckBox ChkPublic;
		private System.Windows.Forms.Label LblTimeRange;
		private System.Windows.Forms.CheckBox ChkStartFrom;
		private System.Windows.Forms.CheckBox ChkStartTo;
		private System.Windows.Forms.Label LblMessages;
		private System.Windows.Forms.ListView LstMessages;
		private System.Windows.Forms.ColumnHeader ClmStatus;
		private System.Windows.Forms.ColumnHeader ClmType;
		private System.Windows.Forms.ColumnHeader ClmSender;
		private System.Windows.Forms.ColumnHeader ClmMessageId;
		private System.Windows.Forms.ColumnHeader ClmBody;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetMemberMessages()
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
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblMessages = new System.Windows.Forms.Label();
			this.LstMessages = new System.Windows.Forms.ListView();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmStatus = new System.Windows.Forms.ColumnHeader();
			this.ClmType = new System.Windows.Forms.ColumnHeader();
			this.ClmSender = new System.Windows.Forms.ColumnHeader();
			this.BtnGetMemberMessages = new System.Windows.Forms.Button();
			this.DatePickStartTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickStartFrom = new System.Windows.Forms.DateTimePicker();
			this.LblStartSep = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.CboStatus = new System.Windows.Forms.ComboBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.CboType = new System.Windows.Forms.ComboBox();
			this.LblType = new System.Windows.Forms.Label();
			this.ChkPublic = new System.Windows.Forms.CheckBox();
			this.LblTimeRange = new System.Windows.Forms.Label();
			this.ChkStartFrom = new System.Windows.Forms.CheckBox();
			this.ChkStartTo = new System.Windows.Forms.CheckBox();
			this.ClmMessageId = new System.Windows.Forms.ColumnHeader();
			this.ClmBody = new System.Windows.Forms.ColumnHeader();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblMessages);
			this.GrpResult.Controls.Add(this.LstMessages);
			this.GrpResult.Location = new System.Drawing.Point(8, 168);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(456, 328);
			this.GrpResult.TabIndex = 24;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Results";
			// 
			// LblMessages
			// 
			this.LblMessages.Location = new System.Drawing.Point(16, 24);
			this.LblMessages.Name = "LblMessages";
			this.LblMessages.Size = new System.Drawing.Size(112, 23);
			this.LblMessages.TabIndex = 15;
			this.LblMessages.Text = "Messages:";
			// 
			// LstMessages
			// 
			this.LstMessages.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.LstMessages.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						  this.ClmItemId,
																						  this.ClmMessageId,
																						  this.ClmStatus,
																						  this.ClmType,
																						  this.ClmSender,
																						  this.ClmBody});
			this.LstMessages.GridLines = true;
			this.LstMessages.Location = new System.Drawing.Point(8, 48);
			this.LstMessages.Name = "LstMessages";
			this.LstMessages.Size = new System.Drawing.Size(440, 264);
			this.LstMessages.TabIndex = 4;
			this.LstMessages.View = System.Windows.Forms.View.Details;
			// 
			// ClmItemId
			// 
			this.ClmItemId.Text = "ItemId";
			this.ClmItemId.Width = 56;
			// 
			// ClmStatus
			// 
			this.ClmStatus.Text = "Status";
			this.ClmStatus.Width = 50;
			// 
			// ClmType
			// 
			this.ClmType.Text = "Question Type";
			this.ClmType.Width = 87;
			// 
			// ClmSender
			// 
			this.ClmSender.Text = "Sender";
			this.ClmSender.Width = 51;
			// 
			// BtnGetMemberMessages
			// 
			this.BtnGetMemberMessages.Location = new System.Drawing.Point(136, 144);
			this.BtnGetMemberMessages.Name = "BtnGetMemberMessages";
			this.BtnGetMemberMessages.Size = new System.Drawing.Size(128, 23);
			this.BtnGetMemberMessages.TabIndex = 23;
			this.BtnGetMemberMessages.Text = "GetMemberMessages";
			this.BtnGetMemberMessages.Click += new System.EventHandler(this.BtnGetMemberMessages_Click);
			// 
			// DatePickStartTo
			// 
			this.DatePickStartTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickStartTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickStartTo.Location = new System.Drawing.Point(288, 104);
			this.DatePickStartTo.Name = "DatePickStartTo";
			this.DatePickStartTo.Size = new System.Drawing.Size(136, 20);
			this.DatePickStartTo.TabIndex = 65;
			// 
			// DatePickStartFrom
			// 
			this.DatePickStartFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickStartFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickStartFrom.Location = new System.Drawing.Point(104, 104);
			this.DatePickStartFrom.Name = "DatePickStartFrom";
			this.DatePickStartFrom.Size = new System.Drawing.Size(136, 20);
			this.DatePickStartFrom.TabIndex = 64;
			// 
			// LblStartSep
			// 
			this.LblStartSep.Location = new System.Drawing.Point(248, 104);
			this.LblStartSep.Name = "LblStartSep";
			this.LblStartSep.Size = new System.Drawing.Size(16, 23);
			this.LblStartSep.TabIndex = 63;
			this.LblStartSep.Text = " - ";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(88, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.TabIndex = 70;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(8, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(80, 16);
			this.LblItemId.TabIndex = 72;
			this.LblItemId.Text = "Item Id:";
			// 
			// CboStatus
			// 
			this.CboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboStatus.Location = new System.Drawing.Point(88, 32);
			this.CboStatus.Name = "CboStatus";
			this.CboStatus.Size = new System.Drawing.Size(144, 21);
			this.CboStatus.TabIndex = 74;
			// 
			// LblStatus
			// 
			this.LblStatus.Location = new System.Drawing.Point(8, 32);
			this.LblStatus.Name = "LblStatus";
			this.LblStatus.Size = new System.Drawing.Size(80, 18);
			this.LblStatus.TabIndex = 73;
			this.LblStatus.Text = "Status:";
			// 
			// CboType
			// 
			this.CboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboType.Location = new System.Drawing.Point(88, 56);
			this.CboType.Name = "CboType";
			this.CboType.Size = new System.Drawing.Size(144, 21);
			this.CboType.TabIndex = 76;
			// 
			// LblType
			// 
			this.LblType.Location = new System.Drawing.Point(8, 56);
			this.LblType.Name = "LblType";
			this.LblType.Size = new System.Drawing.Size(80, 18);
			this.LblType.TabIndex = 75;
			this.LblType.Text = "Type:";
			// 
			// ChkPublic
			// 
			this.ChkPublic.Location = new System.Drawing.Point(88, 80);
			this.ChkPublic.Name = "ChkPublic";
			this.ChkPublic.Size = new System.Drawing.Size(160, 24);
			this.ChkPublic.TabIndex = 77;
			this.ChkPublic.Text = "Retrieve Public Only";
			// 
			// LblTimeRange
			// 
			this.LblTimeRange.Location = new System.Drawing.Point(16, 104);
			this.LblTimeRange.Name = "LblTimeRange";
			this.LblTimeRange.Size = new System.Drawing.Size(64, 16);
			this.LblTimeRange.TabIndex = 78;
			this.LblTimeRange.Text = "Time Filter:";
			// 
			// ChkStartFrom
			// 
			this.ChkStartFrom.Location = new System.Drawing.Point(88, 104);
			this.ChkStartFrom.Name = "ChkStartFrom";
			this.ChkStartFrom.Size = new System.Drawing.Size(12, 24);
			this.ChkStartFrom.TabIndex = 79;
			// 
			// ChkStartTo
			// 
			this.ChkStartTo.Location = new System.Drawing.Point(272, 104);
			this.ChkStartTo.Name = "ChkStartTo";
			this.ChkStartTo.Size = new System.Drawing.Size(12, 24);
			this.ChkStartTo.TabIndex = 80;
			// 
			// ClmMessageId
			// 
			this.ClmMessageId.Text = "Message Id";
			this.ClmMessageId.Width = 69;
			// 
			// ClmBody
			// 
			this.ClmBody.Text = "Question";
			this.ClmBody.Width = 112;
			// 
			// FrmGetMemberMessages
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 501);
			this.Controls.Add(this.ChkStartTo);
			this.Controls.Add(this.ChkStartFrom);
			this.Controls.Add(this.LblTimeRange);
			this.Controls.Add(this.ChkPublic);
			this.Controls.Add(this.CboType);
			this.Controls.Add(this.LblType);
			this.Controls.Add(this.CboStatus);
			this.Controls.Add(this.LblStatus);
			this.Controls.Add(this.LblItemId);
			this.Controls.Add(this.TxtItemId);
			this.Controls.Add(this.DatePickStartTo);
			this.Controls.Add(this.DatePickStartFrom);
			this.Controls.Add(this.LblStartSep);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetMemberMessages);
			this.Name = "FrmGetMemberMessages";
			this.Text = "GetMemberMessages";
			this.Load += new System.EventHandler(this.FrmGetMemberMessages_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

	

		private void FrmGetMemberMessages_Load(object sender, System.EventArgs e)
		{
			DateTime now = DateTime.Now;
			DatePickStartTo.Value = now;
			DatePickStartFrom.Value = now.AddDays(-5);
			
			
			string[] types = Enum.GetNames(typeof(MessageTypeCodeType));
			foreach (string typ in types)
			{
				if (typ != "CustomCode")
				{
					CboType.Items.Add(typ);
				}
			}
			CboType.SelectedIndex = 0;

			string[] statie = Enum.GetNames(typeof(MessageStatusTypeCodeType));
			foreach (string stat in statie)
			{
				if (stat != "CustomCode")
				{
					CboStatus.Items.Add(stat);
				}
			}
			CboStatus.SelectedIndex = 0;

		}


		private void BtnGetMemberMessages_Click(object sender, System.EventArgs e)
		{
			try 
			{
				LstMessages.Items.Clear();

				GetMemberMessagesCall apicall = new GetMemberMessagesCall(Context);				apicall.DisplayToPublic = ChkPublic.Checked;
				MemberMessageExchangeTypeCollection messages;
				if (TxtItemId.Text != String.Empty)				
				{				
					messages = apicall.GetMemberMessages(TxtItemId.Text, (MessageTypeCodeType) Enum.Parse(typeof(MessageTypeCodeType), CboType.SelectedItem.ToString()), (MessageStatusTypeCodeType) Enum.Parse(typeof(MessageStatusTypeCodeType), CboStatus.SelectedItem.ToString()));
				} 
				else
				{
					TimeFilter fltr = new TimeFilter();
					if (ChkStartFrom.Checked)					
					{
						fltr.TimeFrom = DatePickStartFrom.Value;					
					}
					if (ChkStartTo.Checked)					
					{						
						fltr.TimeTo = DatePickStartTo.Value;	
					}					
					messages = apicall.GetMemberMessages(fltr, (MessageTypeCodeType) Enum.Parse(typeof(MessageTypeCodeType), CboType.SelectedItem.ToString()), (MessageStatusTypeCodeType) Enum.Parse(typeof(MessageStatusTypeCodeType), CboStatus.SelectedItem.ToString()));
				}

				foreach (MemberMessageExchangeType msg in messages)
				{
					string[] listparams = new string[6];
					if (msg.Item != null)
						listparams[0] = msg.Item.ItemID;
					else
						listparams[0] = "";
					listparams[1] = msg.Question.MessageID;
					listparams[2] = msg.MessageStatus.ToString();
					listparams[3] = msg.Question.QuestionType.ToString();
					listparams[4] = msg.Question.SenderID;
					listparams[5] = msg.Question.Body;

					ListViewItem vi = new ListViewItem(listparams);
					LstMessages.Items.Add(vi);

				}


			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}
}
