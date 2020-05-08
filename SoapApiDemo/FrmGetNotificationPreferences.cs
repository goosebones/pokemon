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
	/// Summary description for GetCategoriesForm.
	/// </summary>
	public class FrmGetNotificationPreferences : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Button BtnGetNotificationPreferences;
		private System.Windows.Forms.ComboBox CboRole;
		private System.Windows.Forms.Label LblRole;
		private System.Windows.Forms.Label LblDelivery;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.TextBox TxtDelivery;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.Label LblUserPref;
		private System.Windows.Forms.ListView LstPreferences;
		private System.Windows.Forms.ColumnHeader ClmPref;
		private System.Windows.Forms.ColumnHeader ClmStatus;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetNotificationPreferences()
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
			this.BtnGetNotificationPreferences = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblDelivery = new System.Windows.Forms.Label();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtDelivery = new System.Windows.Forms.TextBox();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.LblUserPref = new System.Windows.Forms.Label();
			this.LstPreferences = new System.Windows.Forms.ListView();
			this.ClmPref = new System.Windows.Forms.ColumnHeader();
			this.ClmStatus = new System.Windows.Forms.ColumnHeader();
			this.CboRole = new System.Windows.Forms.ComboBox();
			this.LblRole = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetNotificationPreferences
			// 
			this.BtnGetNotificationPreferences.Location = new System.Drawing.Point(176, 40);
			this.BtnGetNotificationPreferences.Name = "BtnGetNotificationPreferences";
			this.BtnGetNotificationPreferences.Size = new System.Drawing.Size(160, 23);
			this.BtnGetNotificationPreferences.TabIndex = 9;
			this.BtnGetNotificationPreferences.Text = "GetNotificationPreferences";
			this.BtnGetNotificationPreferences.Click += new System.EventHandler(this.BtnGetNotificationPreferences_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblDelivery);
			this.GrpResult.Controls.Add(this.LblStatus);
			this.GrpResult.Controls.Add(this.TxtDelivery);
			this.GrpResult.Controls.Add(this.TxtStatus);
			this.GrpResult.Controls.Add(this.LblUserPref);
			this.GrpResult.Controls.Add(this.LstPreferences);
			this.GrpResult.Location = new System.Drawing.Point(8, 72);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(544, 336);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblDelivery
			// 
			this.LblDelivery.Location = new System.Drawing.Point(192, 24);
			this.LblDelivery.Name = "LblDelivery";
			this.LblDelivery.Size = new System.Drawing.Size(80, 23);
			this.LblDelivery.TabIndex = 77;
			this.LblDelivery.Text = "Delivery URL:";
			// 
			// LblStatus
			// 
			this.LblStatus.Location = new System.Drawing.Point(16, 24);
			this.LblStatus.Name = "LblStatus";
			this.LblStatus.Size = new System.Drawing.Size(48, 23);
			this.LblStatus.TabIndex = 75;
			this.LblStatus.Text = "Status:";
			// 
			// TxtDelivery
			// 
			this.TxtDelivery.Location = new System.Drawing.Point(272, 24);
			this.TxtDelivery.Name = "TxtDelivery";
			this.TxtDelivery.ReadOnly = true;
			this.TxtDelivery.Size = new System.Drawing.Size(256, 20);
			this.TxtDelivery.TabIndex = 73;
			this.TxtDelivery.Text = "";
			// 
			// TxtStatus
			// 
			this.TxtStatus.Location = new System.Drawing.Point(64, 24);
			this.TxtStatus.Name = "TxtStatus";
			this.TxtStatus.ReadOnly = true;
			this.TxtStatus.Size = new System.Drawing.Size(104, 20);
			this.TxtStatus.TabIndex = 71;
			this.TxtStatus.Text = "";
			// 
			// LblUserPref
			// 
			this.LblUserPref.Location = new System.Drawing.Point(16, 56);
			this.LblUserPref.Name = "LblUserPref";
			this.LblUserPref.Size = new System.Drawing.Size(112, 23);
			this.LblUserPref.TabIndex = 12;
			this.LblUserPref.Text = "User Preferences:";
			// 
			// LstPreferences
			// 
			this.LstPreferences.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							 this.ClmPref,
																							 this.ClmStatus});
			this.LstPreferences.GridLines = true;
			this.LstPreferences.Location = new System.Drawing.Point(16, 88);
			this.LstPreferences.Name = "LstPreferences";
			this.LstPreferences.Size = new System.Drawing.Size(520, 240);
			this.LstPreferences.TabIndex = 13;
			this.LstPreferences.View = System.Windows.Forms.View.Details;
			// 
			// ClmPref
			// 
			this.ClmPref.Text = "Preference";
			this.ClmPref.Width = 309;
			// 
			// ClmStatus
			// 
			this.ClmStatus.Text = "Status";
			this.ClmStatus.Width = 190;
			// 
			// CboRole
			// 
			this.CboRole.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboRole.Location = new System.Drawing.Point(200, 8);
			this.CboRole.Name = "CboRole";
			this.CboRole.Size = new System.Drawing.Size(160, 21);
			this.CboRole.TabIndex = 59;
			// 
			// LblRole
			// 
			this.LblRole.Location = new System.Drawing.Point(144, 8);
			this.LblRole.Name = "LblRole";
			this.LblRole.Size = new System.Drawing.Size(56, 18);
			this.LblRole.TabIndex = 58;
			this.LblRole.Text = "Role:";
			// 
			// FrmGetNotificationPreferences
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(560, 413);
			this.Controls.Add(this.CboRole);
			this.Controls.Add(this.LblRole);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetNotificationPreferences);
			this.Name = "FrmGetNotificationPreferences";
			this.Text = "GetNotificationPreferences";
			this.Load += new System.EventHandler(this.FrmGetNotificationPreferences_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetNotificationPreferences_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";
				TxtDelivery.Text = "";
				LstPreferences.Items.Clear();
	
				GetNotificationPreferencesCall apicall = new GetNotificationPreferencesCall(Context);
				
				apicall.GetNotificationPreferences((NotificationRoleCodeType) Enum.Parse(typeof(NotificationRoleCodeType), CboRole.SelectedItem.ToString()));

				
				if (apicall.ApplicationDeliveryPreferences != null)
				{
					TxtStatus.Text = apicall.ApplicationDeliveryPreferences.ApplicationEnable.ToString();
					TxtDelivery.Text = apicall.ApplicationDeliveryPreferences.ApplicationURL;
				}
					
				foreach (NotificationEnableType notify in apicall.UserDeliveryPreferenceList)
				{
					string[] listparams = new string[7];
					listparams[0] = notify.EventType.ToString();
					listparams[1] = notify.EventEnable.ToString();
					ListViewItem vi = new ListViewItem(listparams);
					LstPreferences.Items.Add(vi);

				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void FrmGetNotificationPreferences_Load(object sender, System.EventArgs e)
		{
			string[] roles = Enum.GetNames(typeof(NotificationRoleCodeType));
			foreach (string role in roles)
			{
				if (role != "CustomCode")
				{
					CboRole.Items.Add(role);
				}
			}
			CboRole.SelectedIndex = 0;

		}

	}
	
}
