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
	public class FrmEndItem : System.Windows.Forms.Form
	{
		public ApiContext Context;
		public string mItemID;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.Button BtnEndItem;
		private System.Windows.Forms.ComboBox CboReason;
		private System.Windows.Forms.Label LblReason;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmEndItem()
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
			this.BtnEndItem = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.CboReason = new System.Windows.Forms.ComboBox();
			this.LblReason = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnEndItem
			// 
			this.BtnEndItem.Location = new System.Drawing.Point(128, 64);
			this.BtnEndItem.Name = "BtnEndItem";
			this.BtnEndItem.Size = new System.Drawing.Size(120, 23);
			this.BtnEndItem.TabIndex = 23;
			this.BtnEndItem.Text = "EndItem";
			this.BtnEndItem.Click += new System.EventHandler(this.BtnEndItem_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblStatus,
																					this.TxtStatus});
			this.GrpResult.Location = new System.Drawing.Point(8, 96);
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
			this.TxtItemId.Location = new System.Drawing.Point(152, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(80, 20);
			this.TxtItemId.TabIndex = 26;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(72, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(80, 23);
			this.LblItemId.TabIndex = 27;
			this.LblItemId.Text = "Item Id:";
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
			// FrmEndItem
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 165);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.CboReason,
																		  this.LblReason,
																		  this.TxtItemId,
																		  this.LblItemId,
																		  this.GrpResult,
																		  this.BtnEndItem});
			this.Name = "FrmEndItem";
			this.Text = "EndItem";
			this.Load += new System.EventHandler(this.FrmEndItem_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmEndItem_Load(object sender, System.EventArgs e)
		{
			TxtItemId.Text = mItemID;
			string[] reasons = Enum.GetNames(typeof(EndReasonCodeType));
			foreach (string rsn in reasons)
			{
				if (rsn != "CustomCode")
				{
					CboReason.Items.Add(rsn);
				}
			}
			CboReason.SelectedIndex = 0;

		}
		
		
		private void BtnEndItem_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";

				EndItemCall apicall = new EndItemCall(Context);
				apicall.EndItem(TxtItemId.Text, (EndReasonCodeType) Enum.Parse(typeof(EndReasonCodeType), CboReason.SelectedItem.ToString()));
	
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();

			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
