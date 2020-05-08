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
	public class FrmGetPromotionRules : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.ColumnHeader ClmTitle;
		private System.Windows.Forms.ColumnHeader ClmPrice;
		private System.Windows.Forms.Label LblCategory;
		private System.Windows.Forms.ListView LstCrossPromotions;
		private System.Windows.Forms.Button BtnGetPromotionRules;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.ComboBox CboMethod;
		private System.Windows.Forms.Label LblMethod;
		private System.Windows.Forms.Label LblPromotions;
		private System.Windows.Forms.ColumnHeader ClmPriceType;
		private System.Windows.Forms.ColumnHeader ClmListingType;
		private System.Windows.Forms.Label LblStoreCat;
		private System.Windows.Forms.TextBox TxtStoreCat;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetPromotionRules()
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
			this.ClmPrice = new System.Windows.Forms.ColumnHeader();
			this.ClmPriceType = new System.Windows.Forms.ColumnHeader();
			this.ClmListingType = new System.Windows.Forms.ColumnHeader();
			this.BtnGetPromotionRules = new System.Windows.Forms.Button();
			this.LblCategory = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.CboMethod = new System.Windows.Forms.ComboBox();
			this.LblMethod = new System.Windows.Forms.Label();
			this.LblStoreCat = new System.Windows.Forms.Label();
			this.TxtStoreCat = new System.Windows.Forms.TextBox();
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
																								 this.ClmPrice,
																								 this.ClmPriceType,
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
			// ClmPrice
			// 
			this.ClmPrice.Text = "Price";
			this.ClmPrice.Width = 49;
			// 
			// ClmPriceType
			// 
			this.ClmPriceType.Text = "Price Type";
			this.ClmPriceType.Width = 70;
			// 
			// ClmListingType
			// 
			this.ClmListingType.Text = "Listing Type";
			this.ClmListingType.Width = 96;
			// 
			// BtnGetPromotionRules
			// 
			this.BtnGetPromotionRules.Location = new System.Drawing.Point(152, 88);
			this.BtnGetPromotionRules.Name = "BtnGetPromotionRules";
			this.BtnGetPromotionRules.Size = new System.Drawing.Size(128, 23);
			this.BtnGetPromotionRules.TabIndex = 23;
			this.BtnGetPromotionRules.Text = "GetPromotionRules";
			this.BtnGetPromotionRules.Click += new System.EventHandler(this.BtnGetPromotionRules_Click);
			// 
			// LblCategory
			// 
			this.LblCategory.Location = new System.Drawing.Point(104, 8);
			this.LblCategory.Name = "LblCategory";
			this.LblCategory.Size = new System.Drawing.Size(96, 23);
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
			this.CboMethod.Location = new System.Drawing.Point(200, 56);
			this.CboMethod.Name = "CboMethod";
			this.CboMethod.Size = new System.Drawing.Size(136, 21);
			this.CboMethod.TabIndex = 87;
			// 
			// LblMethod
			// 
			this.LblMethod.Location = new System.Drawing.Point(104, 56);
			this.LblMethod.Name = "LblMethod";
			this.LblMethod.Size = new System.Drawing.Size(88, 18);
			this.LblMethod.TabIndex = 86;
			this.LblMethod.Text = "Method:";
			// 
			// LblStoreCat
			// 
			this.LblStoreCat.Location = new System.Drawing.Point(104, 32);
			this.LblStoreCat.Name = "LblStoreCat";
			this.LblStoreCat.Size = new System.Drawing.Size(96, 23);
			this.LblStoreCat.TabIndex = 89;
			this.LblStoreCat.Text = "Store Category:";
			// 
			// TxtStoreCat
			// 
			this.TxtStoreCat.Location = new System.Drawing.Point(200, 32);
			this.TxtStoreCat.Name = "TxtStoreCat";
			this.TxtStoreCat.TabIndex = 88;
			this.TxtStoreCat.Text = "";
			// 
			// FrmGetPromotionRules
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 437);
			this.Controls.Add(this.LblStoreCat);
			this.Controls.Add(this.TxtStoreCat);
			this.Controls.Add(this.CboMethod);
			this.Controls.Add(this.LblMethod);
			this.Controls.Add(this.LblCategory);
			this.Controls.Add(this.TxtItemId);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetPromotionRules);
			this.Name = "FrmGetPromotionRules";
			this.Text = "GetPromotionRules";
			this.Load += new System.EventHandler(this.FrmGetPromotionRules_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

	

		private void FrmGetPromotionRules_Load(object sender, System.EventArgs e)
		{
			string[] enums = Enum.GetNames(typeof(PromotionMethodCodeType));
			foreach (string item in enums)
			{
				if (item != "CustomCode")
					CboMethod.Items.Add(item);
			}
			CboMethod.SelectedIndex = 0;


		}

		private void BtnGetPromotionRules_Click(object sender, System.EventArgs e)
		{
			try
			{
                //LstCrossPromotions.Items.Clear();
                //GetPromotionRulesCall apicall = new GetPromotionRulesCall(Context);

                //if (TxtItemId.Text != String.Empty)
                //{
                //    apicall.GetPromotionRules(TxtItemId.Text, (PromotionMethodCodeType) Enum.Parse(typeof(PromotionMethodCodeType), CboMethod.SelectedItem.ToString()));
                //}
                //else
                //{
                //    apicall.GetPromotionRules(Convert.ToInt32(TxtStoreCat.Text), (PromotionMethodCodeType) Enum.Parse(typeof(PromotionMethodCodeType), CboMethod.SelectedItem.ToString()));
                //}


                //foreach (PromotedItemType promo in promotions.PromotedItem)
                //{
                //    string[] listparams = new string[5];
                //    listparams[0] = promo.ItemID;
                //    listparams[1] = promo.Title;
                //    listparams[2] = promo.PromotionPrice.Value.ToString();
                //    listparams[3] = promo.PromotionPriceType.ToString();
                //    listparams[4] = promo.ListingType.ToString();

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
