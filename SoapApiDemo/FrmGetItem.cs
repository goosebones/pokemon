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
	/// Summary description for GetItemForm.
	/// </summary>
	public class FrmGetItem : System.Windows.Forms.Form
	{
		public ApiContext Context;
		public string mItemID;
		private System.Windows.Forms.Button BtnGetItem;
		private System.Windows.Forms.GroupBox GrpResult;
		internal System.Windows.Forms.Label LblTitle;
		private System.Windows.Forms.TextBox TxtTitle;
		private System.Windows.Forms.TextBox TxtCurrentPrice;
		internal System.Windows.Forms.Label LblCurrentPrice;
		internal System.Windows.Forms.Label LblBidCount;
		private System.Windows.Forms.TextBox TxtBidCount;
		private System.Windows.Forms.TextBox TxtCategory;
		internal System.Windows.Forms.Label LblCategory;
		private System.Windows.Forms.TextBox TxtCategory2;
		internal System.Windows.Forms.Label LblCategory2;
		private System.Windows.Forms.TextBox TxtStartTime;
		internal System.Windows.Forms.Label LblStartTime;
		private System.Windows.Forms.TextBox TxtBuyItNowPrice;
		internal System.Windows.Forms.Label LblBuyItNowPrice;
		internal System.Windows.Forms.Label LblQuantity;
		private System.Windows.Forms.TextBox TxtQuantity;
		internal System.Windows.Forms.Label LblQuantitySold;
		private System.Windows.Forms.TextBox TxtQuantitySold;
		internal System.Windows.Forms.Label LblEndTime;
		private System.Windows.Forms.TextBox TxtEndTime;
		private System.Windows.Forms.TextBox TxtHighBidder;
		internal System.Windows.Forms.Label LblHighBidder;
		internal System.Windows.Forms.Label LblBestOfferEnabled;
		private System.Windows.Forms.TextBox TxtBestOfferEnabled;
		internal System.Windows.Forms.Label LblBestOfferCount;
		private System.Windows.Forms.TextBox TxtBestOfferCount;
		internal System.Windows.Forms.Label lblApplicationData;
		internal System.Windows.Forms.Label LblPictureURL;
		internal System.Windows.Forms.Label LblPayPalEmailAddress;
		private System.Windows.Forms.TextBox TxtApplicationData;
		private System.Windows.Forms.TextBox TxtPictureURL;
		private System.Windows.Forms.TextBox TxtPayPalEmailAddress;
		private System.Windows.Forms.Button BtnRelistItem;
		internal System.Windows.Forms.Label lblListType;
		private System.Windows.Forms.TextBox TxtListType;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Button BtnReviseItem;
		private System.Windows.Forms.Button BtnEndItem;
		internal System.Windows.Forms.Label LblItemIdResult;
		private System.Windows.Forms.TextBox TxtCategoryId;
		internal System.Windows.Forms.Label LblCategoryId;
		internal System.Windows.Forms.Label LblCategory2ID;
		private System.Windows.Forms.TextBox TxtCategory2Id;
		private System.Windows.Forms.TextBox TxtProductID;
		internal System.Windows.Forms.Label LblProductID;
		internal System.Windows.Forms.Label LblSite;
		private System.Windows.Forms.TextBox TxtSite;
		private ItemType fetchedItem;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetItem()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public ItemType getItem() 
		{
			return fetchedItem;
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
			this.BtnGetItem = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblSite = new System.Windows.Forms.Label();
			this.TxtSite = new System.Windows.Forms.TextBox();
			this.TxtProductID = new System.Windows.Forms.TextBox();
			this.LblProductID = new System.Windows.Forms.Label();
			this.TxtCategoryId = new System.Windows.Forms.TextBox();
			this.LblCategoryId = new System.Windows.Forms.Label();
			this.TxtCategory2Id = new System.Windows.Forms.TextBox();
			this.LblCategory2ID = new System.Windows.Forms.Label();
			this.LblItemIdResult = new System.Windows.Forms.Label();
			this.BtnEndItem = new System.Windows.Forms.Button();
			this.BtnReviseItem = new System.Windows.Forms.Button();
			this.lblApplicationData = new System.Windows.Forms.Label();
			this.TxtApplicationData = new System.Windows.Forms.TextBox();
			this.LblPictureURL = new System.Windows.Forms.Label();
			this.TxtPictureURL = new System.Windows.Forms.TextBox();
			this.lblListType = new System.Windows.Forms.Label();
			this.TxtListType = new System.Windows.Forms.TextBox();
			this.TxtPayPalEmailAddress = new System.Windows.Forms.TextBox();
			this.LblPayPalEmailAddress = new System.Windows.Forms.Label();
			this.LblBestOfferEnabled = new System.Windows.Forms.Label();
			this.TxtBestOfferEnabled = new System.Windows.Forms.TextBox();
			this.LblBestOfferCount = new System.Windows.Forms.Label();
			this.TxtBestOfferCount = new System.Windows.Forms.TextBox();
			this.LblTitle = new System.Windows.Forms.Label();
			this.TxtTitle = new System.Windows.Forms.TextBox();
			this.TxtCurrentPrice = new System.Windows.Forms.TextBox();
			this.LblCurrentPrice = new System.Windows.Forms.Label();
			this.LblBidCount = new System.Windows.Forms.Label();
			this.TxtBidCount = new System.Windows.Forms.TextBox();
			this.TxtCategory = new System.Windows.Forms.TextBox();
			this.LblCategory = new System.Windows.Forms.Label();
			this.TxtCategory2 = new System.Windows.Forms.TextBox();
			this.LblCategory2 = new System.Windows.Forms.Label();
			this.TxtStartTime = new System.Windows.Forms.TextBox();
			this.LblStartTime = new System.Windows.Forms.Label();
			this.TxtBuyItNowPrice = new System.Windows.Forms.TextBox();
			this.LblBuyItNowPrice = new System.Windows.Forms.Label();
			this.LblQuantity = new System.Windows.Forms.Label();
			this.TxtQuantity = new System.Windows.Forms.TextBox();
			this.LblQuantitySold = new System.Windows.Forms.Label();
			this.TxtQuantitySold = new System.Windows.Forms.TextBox();
			this.LblEndTime = new System.Windows.Forms.Label();
			this.TxtEndTime = new System.Windows.Forms.TextBox();
			this.TxtHighBidder = new System.Windows.Forms.TextBox();
			this.LblHighBidder = new System.Windows.Forms.Label();
			this.BtnRelistItem = new System.Windows.Forms.Button();
			this.LblItemId = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetItem
			// 
			this.BtnGetItem.Location = new System.Drawing.Point(192, 40);
			this.BtnGetItem.Name = "BtnGetItem";
			this.BtnGetItem.Size = new System.Drawing.Size(104, 23);
			this.BtnGetItem.TabIndex = 28;
			this.BtnGetItem.Text = "GetItem";
			this.BtnGetItem.Click += new System.EventHandler(this.BtnGetItem_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblSite,
																					this.TxtSite,
																					this.TxtProductID,
																					this.LblProductID,
																					this.TxtCategoryId,
																					this.LblCategoryId,
																					this.TxtCategory2Id,
																					this.LblCategory2ID,
																					this.LblItemIdResult,
																					this.BtnEndItem,
																					this.BtnReviseItem,
																					this.lblApplicationData,
																					this.TxtApplicationData,
																					this.LblPictureURL,
																					this.TxtPictureURL,
																					this.lblListType,
																					this.TxtListType,
																					this.TxtPayPalEmailAddress,
																					this.LblPayPalEmailAddress,
																					this.LblBestOfferEnabled,
																					this.TxtBestOfferEnabled,
																					this.LblBestOfferCount,
																					this.TxtBestOfferCount,
																					this.LblTitle,
																					this.TxtTitle,
																					this.TxtCurrentPrice,
																					this.LblCurrentPrice,
																					this.LblBidCount,
																					this.TxtBidCount,
																					this.TxtCategory,
																					this.LblCategory,
																					this.TxtCategory2,
																					this.LblCategory2,
																					this.TxtStartTime,
																					this.LblStartTime,
																					this.TxtBuyItNowPrice,
																					this.LblBuyItNowPrice,
																					this.LblQuantity,
																					this.TxtQuantity,
																					this.LblQuantitySold,
																					this.TxtQuantitySold,
																					this.LblEndTime,
																					this.TxtEndTime,
																					this.TxtHighBidder,
																					this.LblHighBidder,
																					this.BtnRelistItem});
			this.GrpResult.Location = new System.Drawing.Point(8, 72);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(520, 472);
			this.GrpResult.TabIndex = 41;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblSite
			// 
			this.LblSite.Location = new System.Drawing.Point(16, 344);
			this.LblSite.Name = "LblSite";
			this.LblSite.Size = new System.Drawing.Size(72, 16);
			this.LblSite.TabIndex = 99;
			this.LblSite.Text = "Site:";
			// 
			// TxtSite
			// 
			this.TxtSite.Location = new System.Drawing.Point(120, 344);
			this.TxtSite.Name = "TxtSite";
			this.TxtSite.ReadOnly = true;
			this.TxtSite.Size = new System.Drawing.Size(120, 20);
			this.TxtSite.TabIndex = 100;
			this.TxtSite.Text = "";
			// 
			// TxtProductID
			// 
			this.TxtProductID.Location = new System.Drawing.Point(384, 280);
			this.TxtProductID.Name = "TxtProductID";
			this.TxtProductID.ReadOnly = true;
			this.TxtProductID.Size = new System.Drawing.Size(120, 20);
			this.TxtProductID.TabIndex = 98;
			this.TxtProductID.Text = "";
			// 
			// LblProductID
			// 
			this.LblProductID.Location = new System.Drawing.Point(264, 280);
			this.LblProductID.Name = "LblProductID";
			this.LblProductID.Size = new System.Drawing.Size(120, 16);
			this.LblProductID.TabIndex = 97;
			this.LblProductID.Text = "Product ID";
			// 
			// TxtCategoryId
			// 
			this.TxtCategoryId.Location = new System.Drawing.Point(120, 88);
			this.TxtCategoryId.Name = "TxtCategoryId";
			this.TxtCategoryId.ReadOnly = true;
			this.TxtCategoryId.Size = new System.Drawing.Size(120, 20);
			this.TxtCategoryId.TabIndex = 96;
			this.TxtCategoryId.Text = "";
			// 
			// LblCategoryId
			// 
			this.LblCategoryId.Location = new System.Drawing.Point(16, 88);
			this.LblCategoryId.Name = "LblCategoryId";
			this.LblCategoryId.Size = new System.Drawing.Size(72, 23);
			this.LblCategoryId.TabIndex = 93;
			this.LblCategoryId.Text = "Category ID:";
			// 
			// TxtCategory2Id
			// 
			this.TxtCategory2Id.Location = new System.Drawing.Point(384, 88);
			this.TxtCategory2Id.Name = "TxtCategory2Id";
			this.TxtCategory2Id.ReadOnly = true;
			this.TxtCategory2Id.Size = new System.Drawing.Size(120, 20);
			this.TxtCategory2Id.TabIndex = 95;
			this.TxtCategory2Id.Text = "";
			// 
			// LblCategory2ID
			// 
			this.LblCategory2ID.Location = new System.Drawing.Point(264, 88);
			this.LblCategory2ID.Name = "LblCategory2ID";
			this.LblCategory2ID.Size = new System.Drawing.Size(128, 16);
			this.LblCategory2ID.TabIndex = 94;
			this.LblCategory2ID.Text = "Secondary Category ID:";
			// 
			// LblItemIdResult
			// 
			this.LblItemIdResult.Location = new System.Drawing.Point(16, 400);
			this.LblItemIdResult.Name = "LblItemIdResult";
			this.LblItemIdResult.Size = new System.Drawing.Size(216, 16);
			this.LblItemIdResult.TabIndex = 92;
			// 
			// BtnEndItem
			// 
			this.BtnEndItem.Location = new System.Drawing.Point(240, 400);
			this.BtnEndItem.Name = "BtnEndItem";
			this.BtnEndItem.Size = new System.Drawing.Size(76, 23);
			this.BtnEndItem.TabIndex = 91;
			this.BtnEndItem.Text = "EndItem ...";
			this.BtnEndItem.Visible = false;
			this.BtnEndItem.Click += new System.EventHandler(this.BtnEndItem_Click);
			// 
			// BtnReviseItem
			// 
			this.BtnReviseItem.Location = new System.Drawing.Point(432, 400);
			this.BtnReviseItem.Name = "BtnReviseItem";
			this.BtnReviseItem.Size = new System.Drawing.Size(76, 23);
			this.BtnReviseItem.TabIndex = 90;
			this.BtnReviseItem.Text = "ReviseItem ...";
			this.BtnReviseItem.Visible = false;
			this.BtnReviseItem.Click += new System.EventHandler(this.BtnReviseItem_Click);
			// 
			// lblApplicationData
			// 
			this.lblApplicationData.Location = new System.Drawing.Point(16, 280);
			this.lblApplicationData.Name = "lblApplicationData";
			this.lblApplicationData.Size = new System.Drawing.Size(96, 23);
			this.lblApplicationData.TabIndex = 83;
			this.lblApplicationData.Text = "ApplicationData:";
			// 
			// TxtApplicationData
			// 
			this.TxtApplicationData.Location = new System.Drawing.Point(120, 280);
			this.TxtApplicationData.Name = "TxtApplicationData";
			this.TxtApplicationData.ReadOnly = true;
			this.TxtApplicationData.Size = new System.Drawing.Size(120, 20);
			this.TxtApplicationData.TabIndex = 89;
			this.TxtApplicationData.Text = "";
			// 
			// LblPictureURL
			// 
			this.LblPictureURL.Location = new System.Drawing.Point(16, 312);
			this.LblPictureURL.Name = "LblPictureURL";
			this.LblPictureURL.Size = new System.Drawing.Size(72, 23);
			this.LblPictureURL.TabIndex = 81;
			this.LblPictureURL.Text = "PictureURL:";
			// 
			// TxtPictureURL
			// 
			this.TxtPictureURL.Location = new System.Drawing.Point(120, 312);
			this.TxtPictureURL.Name = "TxtPictureURL";
			this.TxtPictureURL.ReadOnly = true;
			this.TxtPictureURL.Size = new System.Drawing.Size(120, 20);
			this.TxtPictureURL.TabIndex = 85;
			this.TxtPictureURL.Text = "";
			// 
			// lblListType
			// 
			this.lblListType.Location = new System.Drawing.Point(264, 344);
			this.lblListType.Name = "lblListType";
			this.lblListType.Size = new System.Drawing.Size(72, 16);
			this.lblListType.TabIndex = 82;
			this.lblListType.Text = "List Type:";
			// 
			// TxtListType
			// 
			this.TxtListType.Location = new System.Drawing.Point(384, 344);
			this.TxtListType.Name = "TxtListType";
			this.TxtListType.ReadOnly = true;
			this.TxtListType.Size = new System.Drawing.Size(120, 20);
			this.TxtListType.TabIndex = 84;
			this.TxtListType.Text = "";
			// 
			// TxtPayPalEmailAddress
			// 
			this.TxtPayPalEmailAddress.Location = new System.Drawing.Point(384, 312);
			this.TxtPayPalEmailAddress.Name = "TxtPayPalEmailAddress";
			this.TxtPayPalEmailAddress.ReadOnly = true;
			this.TxtPayPalEmailAddress.Size = new System.Drawing.Size(120, 20);
			this.TxtPayPalEmailAddress.TabIndex = 87;
			this.TxtPayPalEmailAddress.Text = "";
			// 
			// LblPayPalEmailAddress
			// 
			this.LblPayPalEmailAddress.Location = new System.Drawing.Point(264, 312);
			this.LblPayPalEmailAddress.Name = "LblPayPalEmailAddress";
			this.LblPayPalEmailAddress.Size = new System.Drawing.Size(120, 16);
			this.LblPayPalEmailAddress.TabIndex = 80;
			this.LblPayPalEmailAddress.Text = "PayPalEmailAddress:";
			// 
			// LblBestOfferEnabled
			// 
			this.LblBestOfferEnabled.Location = new System.Drawing.Point(16, 248);
			this.LblBestOfferEnabled.Name = "LblBestOfferEnabled";
			this.LblBestOfferEnabled.Size = new System.Drawing.Size(104, 23);
			this.LblBestOfferEnabled.TabIndex = 74;
			this.LblBestOfferEnabled.Text = "Best Offer Enabled:";
			// 
			// TxtBestOfferEnabled
			// 
			this.TxtBestOfferEnabled.Location = new System.Drawing.Point(120, 248);
			this.TxtBestOfferEnabled.Name = "TxtBestOfferEnabled";
			this.TxtBestOfferEnabled.ReadOnly = true;
			this.TxtBestOfferEnabled.Size = new System.Drawing.Size(120, 20);
			this.TxtBestOfferEnabled.TabIndex = 77;
			this.TxtBestOfferEnabled.Text = "";
			// 
			// LblBestOfferCount
			// 
			this.LblBestOfferCount.Location = new System.Drawing.Point(264, 248);
			this.LblBestOfferCount.Name = "LblBestOfferCount";
			this.LblBestOfferCount.Size = new System.Drawing.Size(104, 16);
			this.LblBestOfferCount.TabIndex = 75;
			this.LblBestOfferCount.Text = "Best Offer Count:";
			// 
			// TxtBestOfferCount
			// 
			this.TxtBestOfferCount.Location = new System.Drawing.Point(384, 248);
			this.TxtBestOfferCount.Name = "TxtBestOfferCount";
			this.TxtBestOfferCount.ReadOnly = true;
			this.TxtBestOfferCount.Size = new System.Drawing.Size(120, 20);
			this.TxtBestOfferCount.TabIndex = 76;
			this.TxtBestOfferCount.Text = "";
			// 
			// LblTitle
			// 
			this.LblTitle.Location = new System.Drawing.Point(16, 24);
			this.LblTitle.Name = "LblTitle";
			this.LblTitle.Size = new System.Drawing.Size(72, 23);
			this.LblTitle.TabIndex = 61;
			this.LblTitle.Text = "Title:";
			// 
			// TxtTitle
			// 
			this.TxtTitle.Location = new System.Drawing.Point(120, 24);
			this.TxtTitle.Name = "TxtTitle";
			this.TxtTitle.ReadOnly = true;
			this.TxtTitle.Size = new System.Drawing.Size(384, 20);
			this.TxtTitle.TabIndex = 72;
			this.TxtTitle.Text = "";
			this.TxtTitle.TextChanged += new System.EventHandler(this.TxtTitle_TextChanged);
			// 
			// TxtCurrentPrice
			// 
			this.TxtCurrentPrice.Location = new System.Drawing.Point(120, 120);
			this.TxtCurrentPrice.Name = "TxtCurrentPrice";
			this.TxtCurrentPrice.ReadOnly = true;
			this.TxtCurrentPrice.Size = new System.Drawing.Size(120, 20);
			this.TxtCurrentPrice.TabIndex = 73;
			this.TxtCurrentPrice.Text = "";
			// 
			// LblCurrentPrice
			// 
			this.LblCurrentPrice.Location = new System.Drawing.Point(16, 120);
			this.LblCurrentPrice.Name = "LblCurrentPrice";
			this.LblCurrentPrice.Size = new System.Drawing.Size(72, 23);
			this.LblCurrentPrice.TabIndex = 60;
			this.LblCurrentPrice.Text = "Price:";
			// 
			// LblBidCount
			// 
			this.LblBidCount.Location = new System.Drawing.Point(16, 152);
			this.LblBidCount.Name = "LblBidCount";
			this.LblBidCount.Size = new System.Drawing.Size(72, 23);
			this.LblBidCount.TabIndex = 62;
			this.LblBidCount.Text = "BidCount:";
			// 
			// TxtBidCount
			// 
			this.TxtBidCount.Location = new System.Drawing.Point(120, 152);
			this.TxtBidCount.Name = "TxtBidCount";
			this.TxtBidCount.ReadOnly = true;
			this.TxtBidCount.Size = new System.Drawing.Size(120, 20);
			this.TxtBidCount.TabIndex = 69;
			this.TxtBidCount.Text = "";
			// 
			// TxtCategory
			// 
			this.TxtCategory.Location = new System.Drawing.Point(120, 56);
			this.TxtCategory.Name = "TxtCategory";
			this.TxtCategory.ReadOnly = true;
			this.TxtCategory.Size = new System.Drawing.Size(120, 20);
			this.TxtCategory.TabIndex = 70;
			this.TxtCategory.Text = "";
			// 
			// LblCategory
			// 
			this.LblCategory.Location = new System.Drawing.Point(16, 56);
			this.LblCategory.Name = "LblCategory";
			this.LblCategory.Size = new System.Drawing.Size(72, 23);
			this.LblCategory.TabIndex = 54;
			this.LblCategory.Text = "Category:";
			// 
			// TxtCategory2
			// 
			this.TxtCategory2.Location = new System.Drawing.Point(384, 56);
			this.TxtCategory2.Name = "TxtCategory2";
			this.TxtCategory2.ReadOnly = true;
			this.TxtCategory2.Size = new System.Drawing.Size(120, 20);
			this.TxtCategory2.TabIndex = 67;
			this.TxtCategory2.Text = "";
			// 
			// LblCategory2
			// 
			this.LblCategory2.Location = new System.Drawing.Point(264, 56);
			this.LblCategory2.Name = "LblCategory2";
			this.LblCategory2.Size = new System.Drawing.Size(112, 16);
			this.LblCategory2.TabIndex = 55;
			this.LblCategory2.Text = "Secondary Category:";
			// 
			// TxtStartTime
			// 
			this.TxtStartTime.Location = new System.Drawing.Point(120, 184);
			this.TxtStartTime.Name = "TxtStartTime";
			this.TxtStartTime.ReadOnly = true;
			this.TxtStartTime.Size = new System.Drawing.Size(120, 20);
			this.TxtStartTime.TabIndex = 68;
			this.TxtStartTime.Text = "";
			// 
			// LblStartTime
			// 
			this.LblStartTime.Location = new System.Drawing.Point(16, 184);
			this.LblStartTime.Name = "LblStartTime";
			this.LblStartTime.Size = new System.Drawing.Size(72, 23);
			this.LblStartTime.TabIndex = 52;
			this.LblStartTime.Text = "StartTime:";
			// 
			// TxtBuyItNowPrice
			// 
			this.TxtBuyItNowPrice.Location = new System.Drawing.Point(384, 120);
			this.TxtBuyItNowPrice.Name = "TxtBuyItNowPrice";
			this.TxtBuyItNowPrice.ReadOnly = true;
			this.TxtBuyItNowPrice.Size = new System.Drawing.Size(120, 20);
			this.TxtBuyItNowPrice.TabIndex = 71;
			this.TxtBuyItNowPrice.Text = "";
			// 
			// LblBuyItNowPrice
			// 
			this.LblBuyItNowPrice.Location = new System.Drawing.Point(264, 120);
			this.LblBuyItNowPrice.Name = "LblBuyItNowPrice";
			this.LblBuyItNowPrice.Size = new System.Drawing.Size(72, 16);
			this.LblBuyItNowPrice.TabIndex = 53;
			this.LblBuyItNowPrice.Text = "BIN Price:";
			// 
			// LblQuantity
			// 
			this.LblQuantity.Location = new System.Drawing.Point(16, 216);
			this.LblQuantity.Name = "LblQuantity";
			this.LblQuantity.Size = new System.Drawing.Size(72, 23);
			this.LblQuantity.TabIndex = 58;
			this.LblQuantity.Text = "Quantity:";
			// 
			// TxtQuantity
			// 
			this.TxtQuantity.Location = new System.Drawing.Point(120, 216);
			this.TxtQuantity.Name = "TxtQuantity";
			this.TxtQuantity.ReadOnly = true;
			this.TxtQuantity.Size = new System.Drawing.Size(120, 20);
			this.TxtQuantity.TabIndex = 64;
			this.TxtQuantity.Text = "";
			// 
			// LblQuantitySold
			// 
			this.LblQuantitySold.Location = new System.Drawing.Point(264, 216);
			this.LblQuantitySold.Name = "LblQuantitySold";
			this.LblQuantitySold.Size = new System.Drawing.Size(72, 16);
			this.LblQuantitySold.TabIndex = 59;
			this.LblQuantitySold.Text = "QuantitySold:";
			// 
			// TxtQuantitySold
			// 
			this.TxtQuantitySold.Location = new System.Drawing.Point(384, 216);
			this.TxtQuantitySold.Name = "TxtQuantitySold";
			this.TxtQuantitySold.ReadOnly = true;
			this.TxtQuantitySold.Size = new System.Drawing.Size(120, 20);
			this.TxtQuantitySold.TabIndex = 63;
			this.TxtQuantitySold.Text = "";
			// 
			// LblEndTime
			// 
			this.LblEndTime.Location = new System.Drawing.Point(264, 184);
			this.LblEndTime.Name = "LblEndTime";
			this.LblEndTime.Size = new System.Drawing.Size(72, 16);
			this.LblEndTime.TabIndex = 56;
			this.LblEndTime.Text = "EndTime:";
			// 
			// TxtEndTime
			// 
			this.TxtEndTime.Location = new System.Drawing.Point(384, 184);
			this.TxtEndTime.Name = "TxtEndTime";
			this.TxtEndTime.ReadOnly = true;
			this.TxtEndTime.Size = new System.Drawing.Size(120, 20);
			this.TxtEndTime.TabIndex = 65;
			this.TxtEndTime.Text = "";
			// 
			// TxtHighBidder
			// 
			this.TxtHighBidder.Location = new System.Drawing.Point(384, 152);
			this.TxtHighBidder.Name = "TxtHighBidder";
			this.TxtHighBidder.ReadOnly = true;
			this.TxtHighBidder.Size = new System.Drawing.Size(120, 20);
			this.TxtHighBidder.TabIndex = 66;
			this.TxtHighBidder.Text = "";
			// 
			// LblHighBidder
			// 
			this.LblHighBidder.Location = new System.Drawing.Point(264, 152);
			this.LblHighBidder.Name = "LblHighBidder";
			this.LblHighBidder.Size = new System.Drawing.Size(72, 16);
			this.LblHighBidder.TabIndex = 57;
			this.LblHighBidder.Text = "HighBidder:";
			// 
			// BtnRelistItem
			// 
			this.BtnRelistItem.Location = new System.Drawing.Point(336, 400);
			this.BtnRelistItem.Name = "BtnRelistItem";
			this.BtnRelistItem.Size = new System.Drawing.Size(76, 23);
			this.BtnRelistItem.TabIndex = 42;
			this.BtnRelistItem.Text = "RelistItem ...";
			this.BtnRelistItem.Visible = false;
			this.BtnRelistItem.Click += new System.EventHandler(this.BtnReListItem_Click);
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(112, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(80, 23);
			this.LblItemId.TabIndex = 40;
			this.LblItemId.Text = "Item Id:";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(192, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(128, 20);
			this.TxtItemId.TabIndex = 27;
			this.TxtItemId.Text = "";
			// 
			// FrmGetItem
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(544, 553);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.GrpResult,
																		  this.LblItemId,
																		  this.TxtItemId,
																		  this.BtnGetItem});
			this.Name = "FrmGetItem";
			this.Text = "GetItem";
			this.Load += new System.EventHandler(this.FrmGetItem_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void FrmGetItem_Load(object sender, System.EventArgs e)
		{
			TxtItemId.Text = mItemID;
		}

		private void BtnGetItem_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtTitle.Text = "";
				TxtCategory.Text = "";
				TxtCategory2.Text = "";
				TxtCurrentPrice.Text = "";
				TxtBuyItNowPrice.Text = "";
				TxtBidCount.Text = "";
				TxtHighBidder.Text = "";
				TxtStartTime.Text = "";
				TxtEndTime.Text =  "";
				TxtQuantity.Text = "";
				TxtQuantitySold.Text = "";
				LblItemIdResult.Text = "";
				TxtCategory2Id.Text = "";
				TxtCategoryId.Text = "";

				Context.CallRetry = GetCallRetry();
				
				GetItemCall apicall = new GetItemCall(Context);
				apicall.DetailLevelList.Add(DetailLevelCodeType.ReturnAll);
				fetchedItem = apicall.GetItem(TxtItemId.Text);

				// Set values from the item returned
				TxtTitle.Text = fetchedItem.Title;
				TxtCategory.Text = fetchedItem.PrimaryCategory.CategoryName;
				TxtCategoryId.Text = fetchedItem.PrimaryCategory.CategoryID;
					
				if (fetchedItem.SecondaryCategory != null)
				{
					TxtCategory2.Text = fetchedItem.SecondaryCategory.CategoryName;
					TxtCategory2Id.Text = fetchedItem.SecondaryCategory.CategoryID;
				}

				TxtCurrentPrice.Text = fetchedItem.SellingStatus.CurrentPrice.Value.ToString();
				TxtBuyItNowPrice.Text = fetchedItem.BuyItNowPrice.Value.ToString();
				TxtBidCount.Text = fetchedItem.SellingStatus.BidCount.ToString();

				if (fetchedItem.SellingStatus.HighBidder != null)
					TxtHighBidder.Text = fetchedItem.SellingStatus.HighBidder.UserID;

				TxtStartTime.Text = fetchedItem.ListingDetails.StartTime.ToString();
				TxtEndTime.Text =  fetchedItem.ListingDetails.EndTime.ToString();
				TxtQuantity.Text = fetchedItem.Quantity.ToString();
				TxtQuantitySold.Text = fetchedItem.SellingStatus.QuantitySold.ToString();

				TxtBestOfferCount.Text = "0";
				TxtBestOfferEnabled.Text = "False";
				if (fetchedItem.BestOfferDetails != null)
				{
					TxtBestOfferCount.Text = fetchedItem.BestOfferDetails.BestOfferCount.ToString();
					TxtBestOfferEnabled.Text = fetchedItem.BestOfferDetails.BestOfferEnabled.ToString();
				}

				if (fetchedItem.PayPalEmailAddress != null)
					TxtPayPalEmailAddress.Text = fetchedItem.PayPalEmailAddress.ToString();

				if (fetchedItem.ApplicationData != null)
					TxtApplicationData.Text = fetchedItem.ApplicationData.ToString();

				//if (fetchedItem.ProductListingDetails != null)
				//	TxtProductID.Text = fetchedItem.ProductListingDetails.ProductID;

				TxtPictureURL.Text = "";	
				if (fetchedItem.PictureDetails != null)
				{
					StringCollection PictureURLs = fetchedItem.PictureDetails.PictureURL;
					string PictureURL = "";
					for (int i = 0; PictureURLs != null && i < PictureURLs.Count; i++)
					{
						PictureURL += PictureURLs[i] + ",";
					}
					TxtPictureURL.Text = PictureURL;
				} 
				TxtSite.Text = fetchedItem.Site.ToString();

				
				TxtListType.Text = fetchedItem.ListingType.ToString();
				mItemID = fetchedItem.ItemID;
				LblItemIdResult.Text = "Item Id:  " + fetchedItem.ItemID;
				
				if (LblItemIdResult.Text.Length > 0)
				{
					BtnRelistItem.Visible = true;
					BtnReviseItem.Visible = true;
					BtnEndItem.Visible = true;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
		private static CallRetry GetCallRetry() 
		{
				CallRetry retry = new CallRetry();
				retry.DelayTime = 1000;		// 1 second
				retry.MaximumRetries = 2;
				retry.TriggerHttpStatusCodes = new Int32Collection();
				retry.TriggerHttpStatusCodes.Add(502);
				retry.TriggerHttpStatusCodes.Add(404);
				return retry;
		}

		private void BtnReListItem_Click(object sender, System.EventArgs e)
		{
			FrmRelistItem form = new FrmRelistItem();
			form.mItemID = mItemID;
			form.Context = Context;
			form.ShowDialog();
		}

		private void BtnReviseItem_Click(object sender, System.EventArgs e)
		{
			FrmReviseItem form = new FrmReviseItem();
			form.mItemID = mItemID;
			form.Context = Context;
			form.ShowDialog();
		}

		private void BtnEndItem_Click(object sender, System.EventArgs e)
		{
			FrmEndItem form = new FrmEndItem();
			form.mItemID = mItemID;
			form.Context = Context;
			form.ShowDialog();
		}

		private void TxtTitle_TextChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
