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
	public class FrmGeteBayOfficialTime : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Button BtnGeteBayOfficialTime;
		private System.Windows.Forms.Label LblOfficialTime;
		private System.Windows.Forms.DateTimePicker DatePickOfficialTime;
		private System.ComponentModel.Container components = null;

		public FrmGeteBayOfficialTime()
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
			this.BtnGeteBayOfficialTime = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.DatePickOfficialTime = new System.Windows.Forms.DateTimePicker();
			this.LblOfficialTime = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGeteBayOfficialTime
			// 
			this.BtnGeteBayOfficialTime.Location = new System.Drawing.Point(120, 8);
			this.BtnGeteBayOfficialTime.Name = "BtnGeteBayOfficialTime";
			this.BtnGeteBayOfficialTime.Size = new System.Drawing.Size(120, 23);
			this.BtnGeteBayOfficialTime.TabIndex = 23;
			this.BtnGeteBayOfficialTime.Text = "GeteBayOfficialTime";
			this.BtnGeteBayOfficialTime.Click += new System.EventHandler(this.BtnGeteBayOfficialTime_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.DatePickOfficialTime,
																					this.LblOfficialTime});
			this.GrpResult.Location = new System.Drawing.Point(8, 48);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(352, 56);
			this.GrpResult.TabIndex = 25;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// DatePickOfficialTime
			// 
			this.DatePickOfficialTime.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickOfficialTime.Enabled = false;
			this.DatePickOfficialTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickOfficialTime.Location = new System.Drawing.Point(96, 16);
			this.DatePickOfficialTime.Name = "DatePickOfficialTime";
			this.DatePickOfficialTime.Size = new System.Drawing.Size(160, 20);
			this.DatePickOfficialTime.TabIndex = 61;
			// 
			// LblOfficialTime
			// 
			this.LblOfficialTime.Location = new System.Drawing.Point(16, 16);
			this.LblOfficialTime.Name = "LblOfficialTime";
			this.LblOfficialTime.Size = new System.Drawing.Size(80, 16);
			this.LblOfficialTime.TabIndex = 42;
			this.LblOfficialTime.Text = "Official Time:";
			// 
			// FrmGeteBayOfficialTime
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 117);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.GrpResult,
																		  this.BtnGeteBayOfficialTime});
			this.Name = "FrmGeteBayOfficialTime";
			this.Text = "GeteBayOfficialTime";
			this.Load += new System.EventHandler(this.FrmGeteBayOfficialTime_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmGeteBayOfficialTime_Load(object sender, System.EventArgs e)
		{
			DatePickOfficialTime.Value = DateTimePicker.MinDateTime;

		}
		
		
		private void BtnGeteBayOfficialTime_Click(object sender, System.EventArgs e)
		{
			try
			{
				GeteBayOfficialTimeCall apicall = new GeteBayOfficialTimeCall(Context);
				
				DatePickOfficialTime.Value = apicall.GeteBayOfficialTime();
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}


	}
}
