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
	/// Summary description for GetSuggestedCategoriesForm.
	/// </summary>
	public class FrmGetSuggestedCategories : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label LblCategories;
		private System.Windows.Forms.ListView LstCategories;
		private System.Windows.Forms.Button BtnGetSuggestedCategories;
		private System.Windows.Forms.ColumnHeader ClmCatId;
		private System.Windows.Forms.ColumnHeader ClmName;
		private System.Windows.Forms.Label LblQuery;
		private System.Windows.Forms.TextBox TxtQuery;
		private System.Windows.Forms.ColumnHeader ClmPercent;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetSuggestedCategories()
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
			this.BtnGetSuggestedCategories = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblCategories = new System.Windows.Forms.Label();
			this.LstCategories = new System.Windows.Forms.ListView();
			this.ClmCatId = new System.Windows.Forms.ColumnHeader();
			this.ClmName = new System.Windows.Forms.ColumnHeader();
			this.ClmPercent = new System.Windows.Forms.ColumnHeader();
			this.LblQuery = new System.Windows.Forms.Label();
			this.TxtQuery = new System.Windows.Forms.TextBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetSuggestedCategories
			// 
			this.BtnGetSuggestedCategories.Location = new System.Drawing.Point(168, 40);
			this.BtnGetSuggestedCategories.Name = "BtnGetSuggestedCategories";
			this.BtnGetSuggestedCategories.Size = new System.Drawing.Size(176, 23);
			this.BtnGetSuggestedCategories.TabIndex = 9;
			this.BtnGetSuggestedCategories.Text = "GetSuggestedCategories";
			this.BtnGetSuggestedCategories.Click += new System.EventHandler(this.BtnGetSuggestedCategories_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblCategories);
			this.GrpResult.Controls.Add(this.LstCategories);
			this.GrpResult.Location = new System.Drawing.Point(8, 72);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(504, 336);
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
			this.LblCategories.Text = "Suggested Categories:";
			// 
			// LstCategories
			// 
			this.LstCategories.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							this.ClmCatId,
																							this.ClmName,
																							this.ClmPercent});
			this.LstCategories.GridLines = true;
			this.LstCategories.Location = new System.Drawing.Point(16, 48);
			this.LstCategories.Name = "LstCategories";
			this.LstCategories.Size = new System.Drawing.Size(480, 280);
			this.LstCategories.TabIndex = 13;
			this.LstCategories.View = System.Windows.Forms.View.Details;
			// 
			// ClmCatId
			// 
			this.ClmCatId.Text = "Category Id";
			this.ClmCatId.Width = 70;
			// 
			// ClmName
			// 
			this.ClmName.Text = "Name";
			this.ClmName.Width = 297;
			// 
			// ClmPercent
			// 
			this.ClmPercent.Text = "Percent Found";
			this.ClmPercent.Width = 89;
			// 
			// LblQuery
			// 
			this.LblQuery.Location = new System.Drawing.Point(64, 16);
			this.LblQuery.Name = "LblQuery";
			this.LblQuery.Size = new System.Drawing.Size(96, 23);
			this.LblQuery.TabIndex = 13;
			this.LblQuery.Text = "Query:";
			// 
			// TxtQuery
			// 
			this.TxtQuery.Location = new System.Drawing.Point(160, 16);
			this.TxtQuery.Name = "TxtQuery";
			this.TxtQuery.Size = new System.Drawing.Size(288, 20);
			this.TxtQuery.TabIndex = 14;
			this.TxtQuery.Text = "";
			// 
			// FrmGetSuggestedCategories
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(520, 413);
			this.Controls.Add(this.TxtQuery);
			this.Controls.Add(this.LblQuery);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetSuggestedCategories);
			this.Name = "FrmGetSuggestedCategories";
			this.Text = "GetSuggestedCategoriesForm";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetSuggestedCategories_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstCategories.Items.Clear();
	
				GetSuggestedCategoriesCall apicall = new GetSuggestedCategoriesCall(Context);

				SuggestedCategoryTypeCollection cats = apicall.GetSuggestedCategories(TxtQuery.Text);

				if (cats != null)
				{
					foreach (SuggestedCategoryType category in cats)
					{
						string[] listparams = new string[3];
						listparams[0] = category.Category.CategoryID;
						listparams[1] = String.Join(" : ", category.Category.CategoryParentName.ToArray()) + " : " + category.Category.CategoryName;
						listparams[2] =  category.PercentItemFound.ToString();

						ListViewItem vi = new ListViewItem(listparams);
						this.LstCategories.Items.Add(vi);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

	}
}
