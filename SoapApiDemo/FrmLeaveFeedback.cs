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
	/// Summary description for GetUserForm.
	/// </summary>
	public class FrmLeaveFeedback : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.TextBox TxtUserId;
		private System.Windows.Forms.Label LblUserId;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.Button BtnLeaveFeedback;
		private System.Windows.Forms.TextBox TxtTransactionId;
		private System.Windows.Forms.Label LblTransactionId;
		private System.Windows.Forms.TextBox TxtComments;
		private System.Windows.Forms.Label LblComments;
		private System.Windows.Forms.Label LblDuration;
		private System.Windows.Forms.ComboBox CboType;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.Label LblStatus;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmLeaveFeedback()
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
			this.TxtUserId = new System.Windows.Forms.TextBox();
			this.BtnLeaveFeedback = new System.Windows.Forms.Button();
			this.LblUserId = new System.Windows.Forms.Label();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.TxtTransactionId = new System.Windows.Forms.TextBox();
			this.LblTransactionId = new System.Windows.Forms.Label();
			this.TxtComments = new System.Windows.Forms.TextBox();
			this.LblComments = new System.Windows.Forms.Label();
			this.CboType = new System.Windows.Forms.ComboBox();
			this.LblDuration = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// TxtUserId
			// 
			this.TxtUserId.Location = new System.Drawing.Point(88, 32);
			this.TxtUserId.Name = "TxtUserId";
			this.TxtUserId.Size = new System.Drawing.Size(80, 20);
			this.TxtUserId.TabIndex = 22;
			this.TxtUserId.Text = "";
			// 
			// BtnLeaveFeedback
			// 
			this.BtnLeaveFeedback.Location = new System.Drawing.Point(112, 88);
			this.BtnLeaveFeedback.Name = "BtnLeaveFeedback";
			this.BtnLeaveFeedback.Size = new System.Drawing.Size(120, 23);
			this.BtnLeaveFeedback.TabIndex = 23;
			this.BtnLeaveFeedback.Text = "LeaveFeedback";
			this.BtnLeaveFeedback.Click += new System.EventHandler(this.BtnLeaveFeedback_Click);
			// 
			// LblUserId
			// 
			this.LblUserId.Location = new System.Drawing.Point(8, 32);
			this.LblUserId.Name = "LblUserId";
			this.LblUserId.Size = new System.Drawing.Size(80, 23);
			this.LblUserId.TabIndex = 24;
			this.LblUserId.Text = "Target User Id:";
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblStatus);
			this.GrpResult.Controls.Add(this.TxtStatus);
			this.GrpResult.Location = new System.Drawing.Point(8, 120);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(352, 64);
			this.GrpResult.TabIndex = 25;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblStatus
			// 
			this.LblStatus.Location = new System.Drawing.Point(16, 24);
			this.LblStatus.Name = "LblStatus";
			this.LblStatus.Size = new System.Drawing.Size(80, 23);
			this.LblStatus.TabIndex = 42;
			this.LblStatus.Text = "Status:";
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
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(88, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(80, 20);
			this.TxtItemId.TabIndex = 26;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(8, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(80, 23);
			this.LblItemId.TabIndex = 27;
			this.LblItemId.Text = "Item Id:";
			// 
			// TxtTransactionId
			// 
			this.TxtTransactionId.Location = new System.Drawing.Point(280, 8);
			this.TxtTransactionId.Name = "TxtTransactionId";
			this.TxtTransactionId.Size = new System.Drawing.Size(80, 20);
			this.TxtTransactionId.TabIndex = 28;
			this.TxtTransactionId.Text = "";
			// 
			// LblTransactionId
			// 
			this.LblTransactionId.Location = new System.Drawing.Point(200, 8);
			this.LblTransactionId.Name = "LblTransactionId";
			this.LblTransactionId.Size = new System.Drawing.Size(80, 23);
			this.LblTransactionId.TabIndex = 29;
			this.LblTransactionId.Text = "Transaction Id:";
			// 
			// TxtComments
			// 
			this.TxtComments.Location = new System.Drawing.Point(88, 56);
			this.TxtComments.Name = "TxtComments";
			this.TxtComments.Size = new System.Drawing.Size(272, 20);
			this.TxtComments.TabIndex = 30;
			this.TxtComments.Text = "";
			// 
			// LblComments
			// 
			this.LblComments.Location = new System.Drawing.Point(8, 56);
			this.LblComments.Name = "LblComments";
			this.LblComments.Size = new System.Drawing.Size(88, 23);
			this.LblComments.TabIndex = 31;
			this.LblComments.Text = "Comments:";
			// 
			// CboType
			// 
			this.CboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboType.Location = new System.Drawing.Point(280, 32);
			this.CboType.Name = "CboType";
			this.CboType.Size = new System.Drawing.Size(80, 21);
			this.CboType.TabIndex = 55;
			// 
			// LblDuration
			// 
			this.LblDuration.Location = new System.Drawing.Point(200, 32);
			this.LblDuration.Name = "LblDuration";
			this.LblDuration.Size = new System.Drawing.Size(80, 18);
			this.LblDuration.TabIndex = 54;
			this.LblDuration.Text = "Type:";
			// 
			// FrmLeaveFeedback
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 189);
			this.Controls.Add(this.CboType);
			this.Controls.Add(this.LblDuration);
			this.Controls.Add(this.TxtComments);
			this.Controls.Add(this.TxtTransactionId);
			this.Controls.Add(this.TxtItemId);
			this.Controls.Add(this.TxtUserId);
			this.Controls.Add(this.LblComments);
			this.Controls.Add(this.LblTransactionId);
			this.Controls.Add(this.LblItemId);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnLeaveFeedback);
			this.Controls.Add(this.LblUserId);
			this.Name = "FrmLeaveFeedback";
			this.Text = "LeaveFeedback";
			this.Load += new System.EventHandler(this.FrmLeaveFeedback_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void FrmLeaveFeedback_Load(object sender, System.EventArgs e)
		{
			string[] enums = Enum.GetNames(typeof(CommentTypeCodeType));
			foreach (string item in enums)
			{
				if (item != "CustomCode" && item != "Withdrawn")
					CboType.Items.Add(item);
			}

			CboType.SelectedIndex = 0;
		}
		
		
		private void BtnLeaveFeedback_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";
				LeaveFeedbackCall apicall = new LeaveFeedbackCall(Context);

				CommentTypeCodeType type = (CommentTypeCodeType)Enum.Parse(typeof(CommentTypeCodeType), CboType.SelectedItem.ToString());

				apicall.LeaveFeedback(TxtUserId.Text, TxtItemId.Text, TxtTransactionId.Text, type, TxtComments.Text);
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
