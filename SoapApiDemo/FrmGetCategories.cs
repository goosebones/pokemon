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
	public class FrmGetCategories : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label LblCategories;
		private System.Windows.Forms.ListView LstCategories;
		private System.Windows.Forms.Label LblParent;
		private System.Windows.Forms.TextBox TxtParent;
		private System.Windows.Forms.TextBox TxtLevel;
		private System.Windows.Forms.Label LblLevel;
		private System.Windows.Forms.Button BtnGetCategories;
		private System.Windows.Forms.CheckBox ChkViewLeaf;
		private System.Windows.Forms.ColumnHeader ClmCatId;
		private System.Windows.Forms.ColumnHeader ClmLevel;
		private System.Windows.Forms.ColumnHeader ClmName;
		private System.Windows.Forms.ColumnHeader ClmParent;
		private System.Windows.Forms.ColumnHeader ClmLeaf;
		private System.Windows.Forms.ColumnHeader ClmBestOfferEnabled;
		private System.Windows.Forms.ColumnHeader ClmMinimumReservePrice;
		private System.Windows.Forms.CheckBox ChkEnableCompression;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetCategories()
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
			this.BtnGetCategories = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblCategories = new System.Windows.Forms.Label();
			this.LstCategories = new System.Windows.Forms.ListView();
			this.ClmCatId = new System.Windows.Forms.ColumnHeader();
			this.ClmLevel = new System.Windows.Forms.ColumnHeader();
			this.ClmName = new System.Windows.Forms.ColumnHeader();
			this.ClmParent = new System.Windows.Forms.ColumnHeader();
			this.ClmLeaf = new System.Windows.Forms.ColumnHeader();
			this.ClmBestOfferEnabled = new System.Windows.Forms.ColumnHeader();
			this.ClmMinimumReservePrice = new System.Windows.Forms.ColumnHeader();
			this.ChkViewLeaf = new System.Windows.Forms.CheckBox();
			this.LblParent = new System.Windows.Forms.Label();
			this.TxtParent = new System.Windows.Forms.TextBox();
			this.TxtLevel = new System.Windows.Forms.TextBox();
			this.LblLevel = new System.Windows.Forms.Label();
			this.ChkEnableCompression = new System.Windows.Forms.CheckBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetCategories
			// 
			this.BtnGetCategories.Location = new System.Drawing.Point(200, 96);
			this.BtnGetCategories.Name = "BtnGetCategories";
			this.BtnGetCategories.Size = new System.Drawing.Size(128, 23);
			this.BtnGetCategories.TabIndex = 9;
			this.BtnGetCategories.Text = "GetCategories";
			this.BtnGetCategories.Click += new System.EventHandler(this.BtnGetCategories_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblCategories,
																					this.LstCategories});
			this.GrpResult.Location = new System.Drawing.Point(8, 128);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(760, 280);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblCategories
			// 
			this.LblCategories.Location = new System.Drawing.Point(16, 24);
			this.LblCategories.Name = "LblCategories";
			this.LblCategories.Size = new System.Drawing.Size(475, 23);
			this.LblCategories.TabIndex = 12;
			this.LblCategories.Text = "Categories:";
			// 
			// LstCategories
			// 
			this.LstCategories.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							this.ClmCatId,
																							this.ClmLevel,
																							this.ClmName,
																							this.ClmParent,
																							this.ClmLeaf,
																							this.ClmBestOfferEnabled,
																							this.ClmMinimumReservePrice});
			this.LstCategories.GridLines = true;
			this.LstCategories.Location = new System.Drawing.Point(24, 48);
			this.LstCategories.Name = "LstCategories";
			this.LstCategories.Size = new System.Drawing.Size(720, 208);
			this.LstCategories.TabIndex = 13;
			this.LstCategories.View = System.Windows.Forms.View.Details;
			// 
			// ClmCatId
			// 
			this.ClmCatId.Text = "Category Id";
			this.ClmCatId.Width = 83;
			// 
			// ClmLevel
			// 
			this.ClmLevel.Text = "Level";
			this.ClmLevel.Width = 40;
			// 
			// ClmName
			// 
			this.ClmName.Text = "Name";
			this.ClmName.Width = 150;
			// 
			// ClmParent
			// 
			this.ClmParent.Text = "Parent Id";
			// 
			// ClmLeaf
			// 
			this.ClmLeaf.Text = "Leaf";
			this.ClmLeaf.Width = 40;
			// 
			// ClmBestOfferEnabled
			// 
			this.ClmBestOfferEnabled.Text = "BestOfferEnabled";
			this.ClmBestOfferEnabled.Width = 100;
			// 
			// ClmMinimumReservePrice
			// 
			this.ClmMinimumReservePrice.Text = "MinimumReservePrice";
			this.ClmMinimumReservePrice.Width = 120;
			// 
			// ChkViewLeaf
			// 
			this.ChkViewLeaf.Location = new System.Drawing.Point(152, 64);
			this.ChkViewLeaf.Name = "ChkViewLeaf";
			this.ChkViewLeaf.Size = new System.Drawing.Size(160, 24);
			this.ChkViewLeaf.TabIndex = 11;
			this.ChkViewLeaf.Text = "View Leaf Categories Only";
			// 
			// LblParent
			// 
			this.LblParent.Location = new System.Drawing.Point(112, 16);
			this.LblParent.Name = "LblParent";
			this.LblParent.Size = new System.Drawing.Size(96, 23);
			this.LblParent.TabIndex = 13;
			this.LblParent.Text = "Category Parent:";
			// 
			// TxtParent
			// 
			this.TxtParent.Location = new System.Drawing.Point(208, 16);
			this.TxtParent.Name = "TxtParent";
			this.TxtParent.TabIndex = 14;
			this.TxtParent.Text = "";
			// 
			// TxtLevel
			// 
			this.TxtLevel.Location = new System.Drawing.Point(208, 40);
			this.TxtLevel.Name = "TxtLevel";
			this.TxtLevel.Size = new System.Drawing.Size(48, 20);
			this.TxtLevel.TabIndex = 16;
			this.TxtLevel.Text = "";
			// 
			// LblLevel
			// 
			this.LblLevel.Location = new System.Drawing.Point(112, 40);
			this.LblLevel.Name = "LblLevel";
			this.LblLevel.Size = new System.Drawing.Size(96, 23);
			this.LblLevel.TabIndex = 15;
			this.LblLevel.Text = "Level Limit:";
			// 
			// ChkEnableCompression
			// 
			this.ChkEnableCompression.Checked = true;
			this.ChkEnableCompression.CheckState = System.Windows.Forms.CheckState.Checked;
			this.ChkEnableCompression.Location = new System.Drawing.Point(344, 64);
			this.ChkEnableCompression.Name = "ChkEnableCompression";
			this.ChkEnableCompression.Size = new System.Drawing.Size(216, 24);
			this.ChkEnableCompression.TabIndex = 11;
			this.ChkEnableCompression.Text = "Enable HTTP Compression";
			// 
			// FrmGetCategories
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(784, 413);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.TxtLevel,
																		  this.TxtParent,
																		  this.LblLevel,
																		  this.LblParent,
																		  this.ChkViewLeaf,
																		  this.GrpResult,
																		  this.BtnGetCategories,
																		  this.ChkEnableCompression});
			this.Name = "FrmGetCategories";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetCategories_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstCategories.Items.Clear();
	
				GetCategoriesCall apicall = new GetCategoriesCall(Context);

				// Enable HTTP compression to reduce the download size.
				apicall.EnableCompression = this.ChkEnableCompression.Checked;

				apicall.DetailLevelList.Add(DetailLevelCodeType.ReturnAll);
				apicall.ViewAllNodes = !ChkViewLeaf.Checked;
				
				if (TxtLevel.Text.Length > 0)
					apicall.LevelLimit = Convert.ToInt32(TxtLevel.Text);

				if (TxtParent.Text.Length > 0)
				{
					apicall.CategoryParent = new StringCollection();
					apicall.CategoryParent.Add(TxtParent.Text);

				}
				CategoryTypeCollection cats = apicall.GetCategories();
		
				foreach (CategoryType category in cats)
				{
					string[] listparams = new string[8];
					listparams[0] = category.CategoryID;
					listparams[1] = category.CategoryLevel.ToString();
					listparams[2] = category.CategoryName;
					listparams[3] = category.CategoryParentID[0].ToString();
					listparams[4] = category.LeafCategory.ToString();
					listparams[5] = category.BestOfferEnabled.ToString();
					listparams[6] = apicall.ApiResponse.MinimumReservePrice.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					this.LstCategories.Items.Add(vi);

				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
