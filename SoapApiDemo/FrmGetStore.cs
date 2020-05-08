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
	/// Summary description for GetStoreForm.
	/// </summary>
	public class FrmGetStore : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnGetStore;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label LblName;
		private System.Windows.Forms.TextBox TxtName;
		private System.Windows.Forms.TextBox TxtDescription;
		internal System.Windows.Forms.Label LblDescription;
		internal System.Windows.Forms.Label LblSubscription;
		private System.Windows.Forms.TextBox TxtSubscription;
		private System.Windows.Forms.ListView LstStoreCats;
		private System.Windows.Forms.ColumnHeader ClmCatId;
		private System.Windows.Forms.ColumnHeader ClmName;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetStore()
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
			this.BtnGetStore = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblName = new System.Windows.Forms.Label();
			this.TxtName = new System.Windows.Forms.TextBox();
			this.TxtDescription = new System.Windows.Forms.TextBox();
			this.LblDescription = new System.Windows.Forms.Label();
			this.LblSubscription = new System.Windows.Forms.Label();
			this.TxtSubscription = new System.Windows.Forms.TextBox();
			this.LstStoreCats = new System.Windows.Forms.ListView();
			this.ClmCatId = new System.Windows.Forms.ColumnHeader();
			this.ClmName = new System.Windows.Forms.ColumnHeader();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetStore
			// 
			this.BtnGetStore.Location = new System.Drawing.Point(192, 8);
			this.BtnGetStore.Name = "BtnGetStore";
			this.BtnGetStore.Size = new System.Drawing.Size(104, 23);
			this.BtnGetStore.TabIndex = 28;
			this.BtnGetStore.Text = "GetStore";
			this.BtnGetStore.Click += new System.EventHandler(this.BtnGetStore_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LstStoreCats);
			this.GrpResult.Controls.Add(this.LblName);
			this.GrpResult.Controls.Add(this.TxtName);
			this.GrpResult.Controls.Add(this.TxtDescription);
			this.GrpResult.Controls.Add(this.LblDescription);
			this.GrpResult.Controls.Add(this.LblSubscription);
			this.GrpResult.Controls.Add(this.TxtSubscription);
			this.GrpResult.Location = new System.Drawing.Point(8, 40);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(440, 248);
			this.GrpResult.TabIndex = 41;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblName
			// 
			this.LblName.Location = new System.Drawing.Point(8, 24);
			this.LblName.Name = "LblName";
			this.LblName.Size = new System.Drawing.Size(72, 23);
			this.LblName.TabIndex = 61;
			this.LblName.Text = "Name:";
			// 
			// TxtName
			// 
			this.TxtName.Location = new System.Drawing.Point(80, 24);
			this.TxtName.Name = "TxtName";
			this.TxtName.ReadOnly = true;
			this.TxtName.Size = new System.Drawing.Size(120, 20);
			this.TxtName.TabIndex = 72;
			this.TxtName.Text = "";
			// 
			// TxtDescription
			// 
			this.TxtDescription.Location = new System.Drawing.Point(80, 48);
			this.TxtDescription.Multiline = true;
			this.TxtDescription.Name = "TxtDescription";
			this.TxtDescription.ReadOnly = true;
			this.TxtDescription.Size = new System.Drawing.Size(336, 56);
			this.TxtDescription.TabIndex = 70;
			this.TxtDescription.Text = "";
			// 
			// LblDescription
			// 
			this.LblDescription.Location = new System.Drawing.Point(8, 48);
			this.LblDescription.Name = "LblDescription";
			this.LblDescription.Size = new System.Drawing.Size(72, 23);
			this.LblDescription.TabIndex = 54;
			this.LblDescription.Text = "Description:";
			// 
			// LblSubscription
			// 
			this.LblSubscription.Location = new System.Drawing.Point(224, 24);
			this.LblSubscription.Name = "LblSubscription";
			this.LblSubscription.Size = new System.Drawing.Size(72, 23);
			this.LblSubscription.TabIndex = 58;
			this.LblSubscription.Text = "Subscription:";
			// 
			// TxtSubscription
			// 
			this.TxtSubscription.Location = new System.Drawing.Point(296, 24);
			this.TxtSubscription.Name = "TxtSubscription";
			this.TxtSubscription.ReadOnly = true;
			this.TxtSubscription.Size = new System.Drawing.Size(120, 20);
			this.TxtSubscription.TabIndex = 64;
			this.TxtSubscription.Text = "";
			// 
			// LstStoreCats
			// 
			this.LstStoreCats.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.LstStoreCats.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						   this.ClmCatId,
																						   this.ClmName});
			this.LstStoreCats.GridLines = true;
			this.LstStoreCats.Location = new System.Drawing.Point(0, 120);
			this.LstStoreCats.Name = "LstStoreCats";
			this.LstStoreCats.Size = new System.Drawing.Size(440, 120);
			this.LstStoreCats.TabIndex = 73;
			this.LstStoreCats.View = System.Windows.Forms.View.Details;
			// 
			// ClmCatId
			// 
			this.ClmCatId.Text = "Category Id";
			this.ClmCatId.Width = 80;
			// 
			// ClmName
			// 
			this.ClmName.Text = "Name";
			this.ClmName.Width = 346;
			// 
			// FrmGetStore
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(456, 293);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetStore);
			this.Name = "FrmGetStore";
			this.Text = "GetStore";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetStore_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstStoreCats.Items.Clear();
				TxtName.Text = "";
				TxtSubscription.Text = "";
				TxtDescription.Text = "";

				GetStoreCall apicall = new GetStoreCall(Context);
			
				StoreType store = apicall.GetStore();

				TxtName.Text = store.Name;
				TxtSubscription.Text = store.SubscriptionLevel.ToString();
				TxtDescription.Text = store.Description;

				foreach (StoreCustomCategoryType cat in store.CustomCategories)
				{
					string[] listparams = new string[2];
					listparams[0] = cat.CategoryID.ToString();
					listparams[1] = cat.Name;

					ListViewItem vi = new ListViewItem(listparams);
					LstStoreCats.Items.Add(vi);

				}

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
