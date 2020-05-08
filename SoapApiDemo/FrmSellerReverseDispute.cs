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
	public class FrmSellerReverseDispute : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnSellerReverseDispute;
		private System.Windows.Forms.Label LblDisputeId;
		private System.Windows.Forms.TextBox TxtDisputeId;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.ComboBox CboReason;
		private System.Windows.Forms.Label LblReason;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmSellerReverseDispute()
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
			this.BtnSellerReverseDispute = new System.Windows.Forms.Button();
			this.LblDisputeId = new System.Windows.Forms.Label();
			this.CboReason = new System.Windows.Forms.ComboBox();
			this.LblReason = new System.Windows.Forms.Label();
			this.TxtDisputeId = new System.Windows.Forms.TextBox();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnSellerReverseDispute
			// 
			this.BtnSellerReverseDispute.Location = new System.Drawing.Point(128, 64);
			this.BtnSellerReverseDispute.Name = "BtnSellerReverseDispute";
			this.BtnSellerReverseDispute.Size = new System.Drawing.Size(136, 23);
			this.BtnSellerReverseDispute.TabIndex = 56;
			this.BtnSellerReverseDispute.Text = "SellerReverseDispute";
			this.BtnSellerReverseDispute.Click += new System.EventHandler(this.BtnSellerReverseDispute_Click);
			// 
			// LblDisputeId
			// 
			this.LblDisputeId.Location = new System.Drawing.Point(72, 8);
			this.LblDisputeId.Name = "LblDisputeId";
			this.LblDisputeId.Size = new System.Drawing.Size(80, 23);
			this.LblDisputeId.TabIndex = 27;
			this.LblDisputeId.Text = "Dispute Id:";
			// 
			// CboReason
			// 
			this.CboReason.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboReason.Location = new System.Drawing.Point(152, 32);
			this.CboReason.Name = "CboReason";
			this.CboReason.Size = new System.Drawing.Size(144, 21);
			this.CboReason.TabIndex = 55;
			// 
			// LblReason
			// 
			this.LblReason.Location = new System.Drawing.Point(72, 32);
			this.LblReason.Name = "LblReason";
			this.LblReason.Size = new System.Drawing.Size(80, 18);
			this.LblReason.TabIndex = 54;
			this.LblReason.Text = "Reason:";
			// 
			// TxtDisputeId
			// 
			this.TxtDisputeId.Location = new System.Drawing.Point(152, 8);
			this.TxtDisputeId.Name = "TxtDisputeId";
			this.TxtDisputeId.Size = new System.Drawing.Size(80, 20);
			this.TxtDisputeId.TabIndex = 26;
			this.TxtDisputeId.Text = "";
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblStatus);
			this.GrpResult.Controls.Add(this.TxtStatus);
			this.GrpResult.Location = new System.Drawing.Point(8, 96);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(352, 64);
			this.GrpResult.TabIndex = 59;
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
			// FrmSellerReverseDispute
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 165);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.TxtDisputeId);
			this.Controls.Add(this.CboReason);
			this.Controls.Add(this.LblReason);
			this.Controls.Add(this.LblDisputeId);
			this.Controls.Add(this.BtnSellerReverseDispute);
			this.Name = "FrmSellerReverseDispute";
			this.Text = "SellerReverseDispute";
			this.Load += new System.EventHandler(this.FrmSellerReverseDispute_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmSellerReverseDispute_Load(object sender, System.EventArgs e)
		{
			string[] reason = Enum.GetNames(typeof(DisputeResolutionReasonCodeType));
			foreach (string act in reason)
			{
				if (act != "CustomCode")
				{
					CboReason.Items.Add(act);
				}
			}
			CboReason.SelectedIndex = 0;



		}
		
		
		private void BtnSellerReverseDispute_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";

				SellerReverseDisputeCall apicall = new SellerReverseDisputeCall(Context);
				apicall.SellerReverseDispute(TxtDisputeId.Text, (DisputeResolutionReasonCodeType) Enum.Parse(typeof(DisputeResolutionReasonCodeType), CboReason.SelectedItem.ToString()));
	
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();

			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

	
	}
}
