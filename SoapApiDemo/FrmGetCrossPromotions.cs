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
	/// Summary description for FormGetSellerList.
	/// </summary>
	public class FrmGetCrossPromotions : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.ColumnHeader ClmTitle;
		private System.Windows.Forms.ColumnHeader ClmListingType;
		private System.Windows.Forms.Label LblCategory;
		private System.Windows.Forms.ListView LstCrossPromotions;
		private System.Windows.Forms.Button BtnGetCrossPromotions;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.ComboBox CboMethod;
		private System.Windows.Forms.Label LblMethod;
		private System.Windows.Forms.ComboBox CboViewMode;
		private System.Windows.Forms.Label LblViewMode;
		private System.Windows.Forms.Label LblPromotions;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetCrossPromotions()
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
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblPromotions = new System.Windows.Forms.Label();
			this.LstCrossPromotions = new System.Windows.Forms.ListView();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmTitle = new System.Windows.Forms.ColumnHeader();
			this.ClmListingType = new System.Windows.Forms.ColumnHeader();
			this.BtnGetCrossPromotions = new System.Windows.Forms.Button();
			this.LblCategory = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.CboMethod = new System.Windows.Forms.ComboBox();
			this.LblMethod = new System.Windows.Forms.Label();
			this.CboViewMode = new System.Windows.Forms.ComboBox();
			this.LblViewMode = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblPromotions);
			this.GrpResult.Controls.Add(this.LstCrossPromotions);
			this.GrpResult.Location = new System.Drawing.Point(8, 128);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(456, 296);
			this.GrpResult.TabIndex = 24;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Results";
			// 
			// LblPromotions
			// 
			this.LblPromotions.Location = new System.Drawing.Point(16, 24);
			this.LblPromotions.Name = "LblPromotions";
			this.LblPromotions.Size = new System.Drawing.Size(112, 23);
			this.LblPromotions.TabIndex = 15;
			this.LblPromotions.Text = "Promotions:";
			// 
			// LstCrossPromotions
			// 
			this.LstCrossPromotions.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.LstCrossPromotions.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							   this.ClmItemId,
																							   this.ClmTitle,
																							   this.ClmListingType});
			this.LstCrossPromotions.GridLines = true;
			this.LstCrossPromotions.Location = new System.Drawing.Point(8, 48);
			this.LstCrossPromotions.Name = "LstCrossPromotions";
			this.LstCrossPromotions.Size = new System.Drawing.Size(440, 232);
			this.LstCrossPromotions.TabIndex = 4;
			this.LstCrossPromotions.View = System.Windows.Forms.View.Details;
			// 
			// ClmItemId
			// 
			this.ClmItemId.Text = "ItemId";
			this.ClmItemId.Width = 80;
			// 
			// ClmTitle
			// 
			this.ClmTitle.Text = "Title";
			this.ClmTitle.Width = 135;
			// 
			// ClmListingType
			// 
			this.ClmListingType.Text = "Listing Type";
			this.ClmListingType.Width = 96;
			// 
			// BtnGetCrossPromotions
			// 
			this.BtnGetCrossPromotions.Location = new System.Drawing.Point(152, 88);
			this.BtnGetCrossPromotions.Name = "BtnGetCrossPromotions";
			this.BtnGetCrossPromotions.Size = new System.Drawing.Size(128, 23);
			this.BtnGetCrossPromotions.TabIndex = 23;
			this.BtnGetCrossPromotions.Text = "GetCrossPromotions";
			this.BtnGetCrossPromotions.Click += new System.EventHandler(this.BtnGetCrossPromotions_Click);
			// 
			// LblCategory
			// 
			this.LblCategory.Location = new System.Drawing.Point(120, 8);
			this.LblCategory.Name = "LblCategory";
			this.LblCategory.Size = new System.Drawing.Size(80, 23);
			this.LblCategory.TabIndex = 78;
			this.LblCategory.Text = "Item Id:";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(200, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.TabIndex = 77;
			this.TxtItemId.Text = "";
			// 
			// CboMethod
			// 
			this.CboMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboMethod.Location = new System.Drawing.Point(200, 32);
			this.CboMethod.Name = "CboMethod";
			this.CboMethod.Size = new System.Drawing.Size(136, 21);
			this.CboMethod.TabIndex = 87;
			// 
			// LblMethod
			// 
			this.LblMethod.Location = new System.Drawing.Point(120, 32);
			this.LblMethod.Name = "LblMethod";
			this.LblMethod.Size = new System.Drawing.Size(80, 18);
			this.LblMethod.TabIndex = 86;
			this.LblMethod.Text = "Method:";
			// 
			// CboViewMode
			// 
			this.CboViewMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboViewMode.Location = new System.Drawing.Point(200, 56);
			this.CboViewMode.Name = "CboViewMode";
			this.CboViewMode.Size = new System.Drawing.Size(136, 21);
			this.CboViewMode.TabIndex = 89;
			// 
			// LblViewMode
			// 
			this.LblViewMode.Location = new System.Drawing.Point(120, 56);
			this.LblViewMode.Name = "LblViewMode";
			this.LblViewMode.Size = new System.Drawing.Size(80, 18);
			this.LblViewMode.TabIndex = 88;
			this.LblViewMode.Text = "View Mode:";
			// 
			// FrmGetCrossPromotions
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 437);
			this.Controls.Add(this.CboViewMode);
			this.Controls.Add(this.LblViewMode);
			this.Controls.Add(this.CboMethod);
			this.Controls.Add(this.LblMethod);
			this.Controls.Add(this.LblCategory);
			this.Controls.Add(this.TxtItemId);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetCrossPromotions);
			this.Name = "FrmGetCrossPromotions";
			this.Text = "GetCrossPromotions";
			this.Load += new System.EventHandler(this.FrmGetCrossPromotions_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

	

		private void FrmGetCrossPromotions_Load(object sender, System.EventArgs e)
		{
			string[] enums = Enum.GetNames(typeof(PromotionMethodCodeType));
			foreach (string item in enums)
			{
				if (item != "CustomCode")
					CboMethod.Items.Add(item);
			}
			CboMethod.SelectedIndex = 0;

			enums = Enum.GetNames(typeof(TradingRoleCodeType));
			foreach (string item in enums)
			{
				if (item != "CustomCode")
					CboViewMode.Items.Add(item);
			}

			CboViewMode.SelectedIndex = 0;
		}

		private void BtnGetCrossPromotions_Click(object sender, System.EventArgs e)
		{
			try
			{
                //LstCrossPromotions.Items.Clear();
                //GetCrossPromotionsCall apicall = new GetCrossPromotionsCall(Context);
                //apicall.PromotionViewMode = (TradingRoleCodeType)Enum.Parse(typeof(TradingRoleCodeType), CboViewMode.SelectedItem.ToString());
                //CrossPromotionsType promotions = apicall.GetCrossPromotions(TxtItemId.Text, (PromotionMethodCodeType)Enum.Parse(typeof(PromotionMethodCodeType), CboMethod.SelectedItem.ToString()));
				
				
                //foreach (PromotedItemType promo in promotions.PromotedItem)
                //{
                //    string[] listparams = new string[5];
                //    listparams[0] = promo.ItemID;
                //    listparams[1] = promo.Title;
                //    listparams[3] = promo.ListingType.ToString();

                //    ListViewItem vi = new ListViewItem(listparams);
                //    LstCrossPromotions.Items.Add(vi);

                //}

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

	
	}
}
