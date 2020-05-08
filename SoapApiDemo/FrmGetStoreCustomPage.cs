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
	/// Summary description for GetStoreCustomPageForm.
	/// </summary>
	public class FrmGetStoreCustomPage : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnGetStoreCustomPage;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtPageId;
		private System.Windows.Forms.Label LblPageId;
		private System.Windows.Forms.ListView LstCustomPage;
		private System.Windows.Forms.ColumnHeader ClmTitle;
		private System.Windows.Forms.ColumnHeader ClmPageId;
		private System.Windows.Forms.ColumnHeader ClmStatus;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetStoreCustomPage()
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
			this.BtnGetStoreCustomPage = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LstCustomPage = new System.Windows.Forms.ListView();
			this.ClmTitle = new System.Windows.Forms.ColumnHeader();
			this.ClmPageId = new System.Windows.Forms.ColumnHeader();
			this.TxtPageId = new System.Windows.Forms.TextBox();
			this.LblPageId = new System.Windows.Forms.Label();
			this.ClmStatus = new System.Windows.Forms.ColumnHeader();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetStoreCustomPage
			// 
			this.BtnGetStoreCustomPage.Location = new System.Drawing.Point(168, 40);
			this.BtnGetStoreCustomPage.Name = "BtnGetStoreCustomPage";
			this.BtnGetStoreCustomPage.Size = new System.Drawing.Size(128, 23);
			this.BtnGetStoreCustomPage.TabIndex = 28;
			this.BtnGetStoreCustomPage.Text = "GetStoreCustomPage";
			this.BtnGetStoreCustomPage.Click += new System.EventHandler(this.BtnGetStoreCustomPage_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LstCustomPage);
			this.GrpResult.Location = new System.Drawing.Point(8, 88);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(440, 200);
			this.GrpResult.TabIndex = 41;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LstCustomPage
			// 
			this.LstCustomPage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.LstCustomPage.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							this.ClmTitle,
																							this.ClmPageId,
																							this.ClmStatus});
			this.LstCustomPage.GridLines = true;
			this.LstCustomPage.Location = new System.Drawing.Point(0, 24);
			this.LstCustomPage.Name = "LstCustomPage";
			this.LstCustomPage.Size = new System.Drawing.Size(440, 160);
			this.LstCustomPage.TabIndex = 73;
			this.LstCustomPage.View = System.Windows.Forms.View.Details;
			// 
			// ClmTitle
			// 
			this.ClmTitle.Text = "Title";
			this.ClmTitle.Width = 80;
			// 
			// ClmPageId
			// 
			this.ClmPageId.Text = "Page Id";
			this.ClmPageId.Width = 153;
			// 
			// TxtPageId
			// 
			this.TxtPageId.Location = new System.Drawing.Point(208, 8);
			this.TxtPageId.Name = "TxtPageId";
			this.TxtPageId.Size = new System.Drawing.Size(104, 20);
			this.TxtPageId.TabIndex = 43;
			this.TxtPageId.Text = "";
			// 
			// LblPageId
			// 
			this.LblPageId.Location = new System.Drawing.Point(144, 8);
			this.LblPageId.Name = "LblPageId";
			this.LblPageId.Size = new System.Drawing.Size(64, 23);
			this.LblPageId.TabIndex = 42;
			this.LblPageId.Text = "Page Id:";
			// 
			// ClmStatus
			// 
			this.ClmStatus.Text = "Status";
			this.ClmStatus.Width = 195;
			// 
			// FrmGetStoreCustomPage
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(456, 293);
			this.Controls.Add(this.TxtPageId);
			this.Controls.Add(this.LblPageId);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetStoreCustomPage);
			this.Name = "FrmGetStoreCustomPage";
			this.Text = "GetStoreCustomPage";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetStoreCustomPage_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstCustomPage.Items.Clear();

				GetStoreCustomPageCall apicall = new GetStoreCustomPageCall(Context);
			
				if (TxtPageId.Text != String.Empty)
					apicall.PageID = long.Parse(TxtPageId.Text);

				StoreCustomPageTypeCollection pages = apicall.GetStoreCustomPage();

				foreach (StoreCustomPageType page in pages)
				{
					string[] listparams = new string[3];
					listparams[0] = page.Name;
					listparams[1] = page.PageID.ToString();
					listparams[2] = page.Status.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					LstCustomPage.Items.Add(vi);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
