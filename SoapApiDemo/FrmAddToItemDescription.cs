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
	public class FrmAddToItemDescription : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.Label LblAppend;
		private System.Windows.Forms.TextBox TxtAppend;
		private System.Windows.Forms.Button BtnAddToItemDescriptionCall;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmAddToItemDescription()
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
			this.BtnAddToItemDescriptionCall = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.LblAppend = new System.Windows.Forms.Label();
			this.TxtAppend = new System.Windows.Forms.TextBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnAddToItemDescriptionCall
			// 
			this.BtnAddToItemDescriptionCall.Location = new System.Drawing.Point(112, 104);
			this.BtnAddToItemDescriptionCall.Name = "BtnAddToItemDescriptionCall";
			this.BtnAddToItemDescriptionCall.Size = new System.Drawing.Size(144, 23);
			this.BtnAddToItemDescriptionCall.TabIndex = 23;
			this.BtnAddToItemDescriptionCall.Text = "AddToItemDescriptionCall";
			this.BtnAddToItemDescriptionCall.Click += new System.EventHandler(this.BtnAddToItemDescription_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblStatus);
			this.GrpResult.Controls.Add(this.TxtStatus);
			this.GrpResult.Location = new System.Drawing.Point(8, 136);
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
			// LblAppend
			// 
			this.LblAppend.Location = new System.Drawing.Point(8, 32);
			this.LblAppend.Name = "LblAppend";
			this.LblAppend.Size = new System.Drawing.Size(88, 23);
			this.LblAppend.TabIndex = 31;
			this.LblAppend.Text = "Comments:";
			// 
			// TxtAppend
			// 
			this.TxtAppend.Location = new System.Drawing.Point(88, 32);
			this.TxtAppend.Multiline = true;
			this.TxtAppend.Name = "TxtAppend";
			this.TxtAppend.Size = new System.Drawing.Size(272, 64);
			this.TxtAppend.TabIndex = 30;
			this.TxtAppend.Text = "Append to Description";
			// 
			// FrmAddToItemDescription
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 205);
			this.Controls.Add(this.TxtAppend);
			this.Controls.Add(this.LblAppend);
			this.Controls.Add(this.TxtItemId);
			this.Controls.Add(this.LblItemId);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnAddToItemDescriptionCall);
			this.Name = "FrmAddToItemDescription";
			this.Text = "AddToItemDescription";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

			
		private void BtnAddToItemDescription_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";

				AddToItemDescriptionCall apicall = new AddToItemDescriptionCall(Context);

				apicall.AddToItemDescription(TxtItemId.Text, TxtAppend.Text);
	
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
