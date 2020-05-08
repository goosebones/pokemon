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
	/// Summary description for GetStoreOptionsForm.
	/// </summary>
	public class FrmGetStoreOptions : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnGetStoreOptions;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label LblBasic;
		private System.Windows.Forms.Label LblAdvanced;
		private System.Windows.Forms.Label LblLogos;
		private System.Windows.Forms.Label LblSubscriptions;
		private System.Windows.Forms.ListView LstBasicThemes;
		private System.Windows.Forms.ColumnHeader ClmThemId;
		private System.Windows.Forms.ColumnHeader ClmName;
		private System.Windows.Forms.ListView LstAdvancedThemes;
		private System.Windows.Forms.ListView LstLogos;
		private System.Windows.Forms.ColumnHeader ClmLogoId;
		private System.Windows.Forms.ListView LstSubscriptions;
		private System.Windows.Forms.ColumnHeader ClmLevel;
		private System.Windows.Forms.ColumnHeader ClmFee;
		private System.Windows.Forms.ColumnHeader ClmThemIdA;
		private System.Windows.Forms.ColumnHeader ClmNameA;
		private System.Windows.Forms.ColumnHeader ClmLogoName;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetStoreOptions()
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
			this.BtnGetStoreOptions = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblBasic = new System.Windows.Forms.Label();
			this.LblAdvanced = new System.Windows.Forms.Label();
			this.LblLogos = new System.Windows.Forms.Label();
			this.LblSubscriptions = new System.Windows.Forms.Label();
			this.LstBasicThemes = new System.Windows.Forms.ListView();
			this.ClmThemId = new System.Windows.Forms.ColumnHeader();
			this.ClmName = new System.Windows.Forms.ColumnHeader();
			this.LstAdvancedThemes = new System.Windows.Forms.ListView();
			this.LstLogos = new System.Windows.Forms.ListView();
			this.ClmLogoId = new System.Windows.Forms.ColumnHeader();
			this.LstSubscriptions = new System.Windows.Forms.ListView();
			this.ClmLevel = new System.Windows.Forms.ColumnHeader();
			this.ClmFee = new System.Windows.Forms.ColumnHeader();
			this.ClmThemIdA = new System.Windows.Forms.ColumnHeader();
			this.ClmNameA = new System.Windows.Forms.ColumnHeader();
			this.ClmLogoName = new System.Windows.Forms.ColumnHeader();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetStoreOptions
			// 
			this.BtnGetStoreOptions.Location = new System.Drawing.Point(168, 16);
			this.BtnGetStoreOptions.Name = "BtnGetStoreOptions";
			this.BtnGetStoreOptions.Size = new System.Drawing.Size(128, 23);
			this.BtnGetStoreOptions.TabIndex = 28;
			this.BtnGetStoreOptions.Text = "GetStoreOptions";
			this.BtnGetStoreOptions.Click += new System.EventHandler(this.BtnGetStoreOptions_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LstSubscriptions);
			this.GrpResult.Controls.Add(this.LstLogos);
			this.GrpResult.Controls.Add(this.LstAdvancedThemes);
			this.GrpResult.Controls.Add(this.LstBasicThemes);
			this.GrpResult.Controls.Add(this.LblSubscriptions);
			this.GrpResult.Controls.Add(this.LblLogos);
			this.GrpResult.Controls.Add(this.LblAdvanced);
			this.GrpResult.Controls.Add(this.LblBasic);
			this.GrpResult.Location = new System.Drawing.Point(8, 48);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(440, 568);
			this.GrpResult.TabIndex = 41;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblBasic
			// 
			this.LblBasic.Location = new System.Drawing.Point(8, 16);
			this.LblBasic.Name = "LblBasic";
			this.LblBasic.Size = new System.Drawing.Size(64, 23);
			this.LblBasic.TabIndex = 74;
			this.LblBasic.Text = "Basic";
			// 
			// LblAdvanced
			// 
			this.LblAdvanced.Location = new System.Drawing.Point(8, 152);
			this.LblAdvanced.Name = "LblAdvanced";
			this.LblAdvanced.Size = new System.Drawing.Size(64, 23);
			this.LblAdvanced.TabIndex = 76;
			this.LblAdvanced.Text = "Advanced";
			// 
			// LblLogos
			// 
			this.LblLogos.Location = new System.Drawing.Point(8, 288);
			this.LblLogos.Name = "LblLogos";
			this.LblLogos.Size = new System.Drawing.Size(64, 23);
			this.LblLogos.TabIndex = 78;
			this.LblLogos.Text = "Logos";
			// 
			// LblSubscriptions
			// 
			this.LblSubscriptions.Location = new System.Drawing.Point(8, 424);
			this.LblSubscriptions.Name = "LblSubscriptions";
			this.LblSubscriptions.Size = new System.Drawing.Size(144, 23);
			this.LblSubscriptions.TabIndex = 80;
			this.LblSubscriptions.Text = "Subscriptions";
			// 
			// LstBasicThemes
			// 
			this.LstBasicThemes.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							 this.ClmThemId,
																							 this.ClmName});
			this.LstBasicThemes.GridLines = true;
			this.LstBasicThemes.Location = new System.Drawing.Point(8, 40);
			this.LstBasicThemes.Name = "LstBasicThemes";
			this.LstBasicThemes.Size = new System.Drawing.Size(424, 104);
			this.LstBasicThemes.TabIndex = 81;
			this.LstBasicThemes.View = System.Windows.Forms.View.Details;
			// 
			// ClmThemId
			// 
			this.ClmThemId.Text = "Theme Id";
			this.ClmThemId.Width = 87;
			// 
			// ClmName
			// 
			this.ClmName.Text = "Name";
			this.ClmName.Width = 313;
			// 
			// LstAdvancedThemes
			// 
			this.LstAdvancedThemes.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																								this.ClmThemIdA,
																								this.ClmNameA});
			this.LstAdvancedThemes.GridLines = true;
			this.LstAdvancedThemes.Location = new System.Drawing.Point(8, 176);
			this.LstAdvancedThemes.Name = "LstAdvancedThemes";
			this.LstAdvancedThemes.Size = new System.Drawing.Size(424, 104);
			this.LstAdvancedThemes.TabIndex = 82;
			this.LstAdvancedThemes.View = System.Windows.Forms.View.Details;
			// 
			// LstLogos
			// 
			this.LstLogos.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																					   this.ClmLogoId,
																					   this.ClmLogoName});
			this.LstLogos.GridLines = true;
			this.LstLogos.Location = new System.Drawing.Point(8, 312);
			this.LstLogos.Name = "LstLogos";
			this.LstLogos.Size = new System.Drawing.Size(424, 104);
			this.LstLogos.TabIndex = 83;
			this.LstLogos.View = System.Windows.Forms.View.Details;
			// 
			// ClmLogoId
			// 
			this.ClmLogoId.Text = "Logo Id";
			this.ClmLogoId.Width = 91;
			// 
			// LstSubscriptions
			// 
			this.LstSubscriptions.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							   this.ClmLevel,
																							   this.ClmFee});
			this.LstSubscriptions.GridLines = true;
			this.LstSubscriptions.Location = new System.Drawing.Point(8, 448);
			this.LstSubscriptions.Name = "LstSubscriptions";
			this.LstSubscriptions.Size = new System.Drawing.Size(424, 104);
			this.LstSubscriptions.TabIndex = 84;
			this.LstSubscriptions.View = System.Windows.Forms.View.Details;
			// 
			// ClmLevel
			// 
			this.ClmLevel.Text = "Level";
			this.ClmLevel.Width = 134;
			// 
			// ClmFee
			// 
			this.ClmFee.Text = "Fee";
			this.ClmFee.Width = 270;
			// 
			// ClmThemIdA
			// 
			this.ClmThemIdA.Text = "Theme Id";
			this.ClmThemIdA.Width = 89;
			// 
			// ClmNameA
			// 
			this.ClmNameA.Text = "Name";
			this.ClmNameA.Width = 312;
			// 
			// ClmLogoName
			// 
			this.ClmLogoName.Text = "Name";
			this.ClmLogoName.Width = 312;
			// 
			// FrmGetStoreOptions
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(456, 629);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetStoreOptions);
			this.Name = "FrmGetStoreOptions";
			this.Text = "GetStoreOptions";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetStoreOptions_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstBasicThemes.Items.Clear();
				LstAdvancedThemes.Items.Clear();
				LstLogos.Items.Clear();
				LstSubscriptions.Items.Clear();

				GetStoreOptionsCall apicall = new GetStoreOptionsCall(Context);
				apicall.GetStoreOptions();

				foreach (StoreThemeType theme in apicall.BasicThemeArray.Theme)
				{
					string[] listparams = new string[2];
					listparams[0] = theme.ThemeID.ToString();
					listparams[1] = theme.Name;

					ListViewItem vi = new ListViewItem(listparams);
					LstBasicThemes.Items.Add(vi);
				}

				foreach (StoreThemeType theme in apicall.AdvancedThemeArray.Theme)
				{
					string[] listparams = new string[2];
					listparams[0] = theme.ThemeID.ToString();
					listparams[1] = theme.Name;

					ListViewItem vi = new ListViewItem(listparams);
					LstAdvancedThemes.Items.Add(vi);
				}
				foreach (StoreLogoType logo in apicall.LogoList)
				{
					string[] listparams = new string[2];
					listparams[0] = logo.LogoID.ToString();
					listparams[1] = logo.Name;

					ListViewItem vi = new ListViewItem(listparams);
					LstLogos.Items.Add(vi);
				}
				foreach (StoreSubscriptionType subs in apicall.SubscriptionList)
				{
					string[] listparams = new string[2];
					listparams[0] = subs.Level.ToString();
					listparams[1] = subs.Fee.Value.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					LstSubscriptions.Items.Add(vi);
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
