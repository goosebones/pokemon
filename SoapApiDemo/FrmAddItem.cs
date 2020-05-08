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
	/// Summary description for AddItemForm.
	/// </summary>
	public class FrmAddItem : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox grpResults;
		private System.Windows.Forms.OpenFileDialog OpenFileDialogIMG;
		private System.Windows.Forms.TextBox TxtStartPrice;
		private System.Windows.Forms.TextBox TxtTitle;
		private System.Windows.Forms.TextBox TxtDescription;
		private System.Windows.Forms.TextBox TxtBuyItNowPrice;
		private System.Windows.Forms.TextBox TxtCategory;
		private System.Windows.Forms.Label LblCategory;
		private System.Windows.Forms.Label LblStartPrice;
		private System.Windows.Forms.Label LblDescription;
		private System.Windows.Forms.Label LblTitle;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.TextBox TxtListingFee;
		private System.Windows.Forms.Label LblListingFee;
		private System.Windows.Forms.TextBox TxtReservePrice;
		private System.Windows.Forms.Label LblBuyItNowPrice;
		private System.Windows.Forms.Button BtnAddItem;
		private System.Windows.Forms.Label LblReservePrice;
		private System.Windows.Forms.GroupBox GrpPictures;
		private System.Windows.Forms.Button BtnAddPic;
		private System.Windows.Forms.ListBox ListPictures;
		private System.Windows.Forms.Button BtnRemovePic;
		private System.Windows.Forms.CheckBox ChkBoldTitle;
		private System.Windows.Forms.CheckBox ChkHighLight;
		private System.Windows.Forms.TextBox TxtQuantity;
		private System.Windows.Forms.Label LblQuantity;
		private System.Windows.Forms.Label LblDuration;
		private System.Windows.Forms.TextBox TxtLocation;
		private System.Windows.Forms.Label LblLocation;
		private System.Windows.Forms.Button BtnVerifyAddItem;
		private System.Windows.Forms.ComboBox CboDuration;
		private System.Windows.Forms.ComboBox CboPicDisplay;
		private System.Windows.Forms.CheckBox ChkEnableBestOffer;
		private System.Windows.Forms.TextBox TxtCategory2;
		private System.Windows.Forms.Label lblCategory2;
		private System.Windows.Forms.TextBox TxtApplicationData;
		private System.Windows.Forms.Label lblApplicationData;
		private System.Windows.Forms.Label LblPayPalEmailAddress;
		private System.Windows.Forms.TextBox TxtPayPalEmailAddress;
		private System.Windows.Forms.ComboBox CboListType;
		private System.Windows.Forms.Label LblListType;
		private System.Windows.Forms.Button BtnGetItem;
		private System.Windows.Forms.Button BtnGetCategories;
        private Label conditionLabel;
        private ComboBox CboCondition;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmAddItem()
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
            this.TxtStartPrice = new System.Windows.Forms.TextBox();
            this.TxtTitle = new System.Windows.Forms.TextBox();
            this.TxtDescription = new System.Windows.Forms.TextBox();
            this.TxtBuyItNowPrice = new System.Windows.Forms.TextBox();
            this.TxtCategory = new System.Windows.Forms.TextBox();
            this.LblCategory = new System.Windows.Forms.Label();
            this.LblStartPrice = new System.Windows.Forms.Label();
            this.LblDescription = new System.Windows.Forms.Label();
            this.LblTitle = new System.Windows.Forms.Label();
            this.grpResults = new System.Windows.Forms.GroupBox();
            this.BtnGetItem = new System.Windows.Forms.Button();
            this.LblItemId = new System.Windows.Forms.Label();
            this.TxtItemId = new System.Windows.Forms.TextBox();
            this.TxtListingFee = new System.Windows.Forms.TextBox();
            this.LblListingFee = new System.Windows.Forms.Label();
            this.TxtReservePrice = new System.Windows.Forms.TextBox();
            this.LblBuyItNowPrice = new System.Windows.Forms.Label();
            this.BtnAddItem = new System.Windows.Forms.Button();
            this.OpenFileDialogIMG = new System.Windows.Forms.OpenFileDialog();
            this.LblReservePrice = new System.Windows.Forms.Label();
            this.GrpPictures = new System.Windows.Forms.GroupBox();
            this.CboPicDisplay = new System.Windows.Forms.ComboBox();
            this.BtnAddPic = new System.Windows.Forms.Button();
            this.ListPictures = new System.Windows.Forms.ListBox();
            this.BtnRemovePic = new System.Windows.Forms.Button();
            this.ChkBoldTitle = new System.Windows.Forms.CheckBox();
            this.ChkHighLight = new System.Windows.Forms.CheckBox();
            this.TxtQuantity = new System.Windows.Forms.TextBox();
            this.LblQuantity = new System.Windows.Forms.Label();
            this.LblDuration = new System.Windows.Forms.Label();
            this.TxtLocation = new System.Windows.Forms.TextBox();
            this.LblLocation = new System.Windows.Forms.Label();
            this.BtnVerifyAddItem = new System.Windows.Forms.Button();
            this.CboDuration = new System.Windows.Forms.ComboBox();
            this.ChkEnableBestOffer = new System.Windows.Forms.CheckBox();
            this.TxtCategory2 = new System.Windows.Forms.TextBox();
            this.lblCategory2 = new System.Windows.Forms.Label();
            this.TxtApplicationData = new System.Windows.Forms.TextBox();
            this.lblApplicationData = new System.Windows.Forms.Label();
            this.TxtPayPalEmailAddress = new System.Windows.Forms.TextBox();
            this.LblPayPalEmailAddress = new System.Windows.Forms.Label();
            this.CboListType = new System.Windows.Forms.ComboBox();
            this.LblListType = new System.Windows.Forms.Label();
            this.BtnGetCategories = new System.Windows.Forms.Button();
            this.conditionLabel = new System.Windows.Forms.Label();
            this.CboCondition = new System.Windows.Forms.ComboBox();
            this.grpResults.SuspendLayout();
            this.GrpPictures.SuspendLayout();
            this.SuspendLayout();
            // 
            // TxtStartPrice
            // 
            this.TxtStartPrice.Location = new System.Drawing.Point(80, 160);
            this.TxtStartPrice.Name = "TxtStartPrice";
            this.TxtStartPrice.Size = new System.Drawing.Size(100, 20);
            this.TxtStartPrice.TabIndex = 33;
            this.TxtStartPrice.Text = "1.00";
            // 
            // TxtTitle
            // 
            this.TxtTitle.Location = new System.Drawing.Point(72, 8);
            this.TxtTitle.Name = "TxtTitle";
            this.TxtTitle.Size = new System.Drawing.Size(100, 20);
            this.TxtTitle.TabIndex = 30;
            this.TxtTitle.Text = "My ItemTitle";
            // 
            // TxtDescription
            // 
            this.TxtDescription.Location = new System.Drawing.Point(104, 224);
            this.TxtDescription.Multiline = true;
            this.TxtDescription.Name = "TxtDescription";
            this.TxtDescription.Size = new System.Drawing.Size(100, 20);
            this.TxtDescription.TabIndex = 35;
            this.TxtDescription.Text = "My item description.";
            // 
            // TxtBuyItNowPrice
            // 
            this.TxtBuyItNowPrice.Location = new System.Drawing.Point(456, 160);
            this.TxtBuyItNowPrice.Name = "TxtBuyItNowPrice";
            this.TxtBuyItNowPrice.Size = new System.Drawing.Size(100, 20);
            this.TxtBuyItNowPrice.TabIndex = 34;
            this.TxtBuyItNowPrice.Text = "10.0";
            // 
            // TxtCategory
            // 
            this.TxtCategory.Location = new System.Drawing.Point(72, 96);
            this.TxtCategory.Name = "TxtCategory";
            this.TxtCategory.Size = new System.Drawing.Size(100, 20);
            this.TxtCategory.TabIndex = 31;
            this.TxtCategory.Text = "11104";
            // 
            // LblCategory
            // 
            this.LblCategory.Location = new System.Drawing.Point(8, 96);
            this.LblCategory.Name = "LblCategory";
            this.LblCategory.Size = new System.Drawing.Size(100, 23);
            this.LblCategory.TabIndex = 38;
            this.LblCategory.Text = "Category:";
            // 
            // LblStartPrice
            // 
            this.LblStartPrice.Location = new System.Drawing.Point(8, 160);
            this.LblStartPrice.Name = "LblStartPrice";
            this.LblStartPrice.Size = new System.Drawing.Size(100, 23);
            this.LblStartPrice.TabIndex = 40;
            this.LblStartPrice.Text = "Start Price:";
            // 
            // LblDescription
            // 
            this.LblDescription.Location = new System.Drawing.Point(8, 224);
            this.LblDescription.Name = "LblDescription";
            this.LblDescription.Size = new System.Drawing.Size(100, 23);
            this.LblDescription.TabIndex = 42;
            this.LblDescription.Text = "Description:";
            // 
            // LblTitle
            // 
            this.LblTitle.Location = new System.Drawing.Point(8, 8);
            this.LblTitle.Name = "LblTitle";
            this.LblTitle.Size = new System.Drawing.Size(100, 23);
            this.LblTitle.TabIndex = 37;
            this.LblTitle.Text = "Title:";
            // 
            // grpResults
            // 
            this.grpResults.Controls.Add(this.BtnGetItem);
            this.grpResults.Controls.Add(this.LblItemId);
            this.grpResults.Controls.Add(this.TxtItemId);
            this.grpResults.Controls.Add(this.TxtListingFee);
            this.grpResults.Controls.Add(this.LblListingFee);
            this.grpResults.Location = new System.Drawing.Point(24, 496);
            this.grpResults.Name = "grpResults";
            this.grpResults.Size = new System.Drawing.Size(552, 100);
            this.grpResults.TabIndex = 44;
            this.grpResults.TabStop = false;
            this.grpResults.Text = "Results";
            // 
            // BtnGetItem
            // 
            this.BtnGetItem.Location = new System.Drawing.Point(208, 24);
            this.BtnGetItem.Name = "BtnGetItem";
            this.BtnGetItem.Size = new System.Drawing.Size(75, 23);
            this.BtnGetItem.TabIndex = 2;
            this.BtnGetItem.Text = "GetItem...";
            this.BtnGetItem.Visible = false;
            this.BtnGetItem.Click += new System.EventHandler(this.BtnGetItem_Click);
            // 
            // LblItemId
            // 
            this.LblItemId.Location = new System.Drawing.Point(8, 24);
            this.LblItemId.Name = "LblItemId";
            this.LblItemId.Size = new System.Drawing.Size(100, 23);
            this.LblItemId.TabIndex = 1;
            this.LblItemId.Text = "ItemId:";
            // 
            // TxtItemId
            // 
            this.TxtItemId.Location = new System.Drawing.Point(104, 24);
            this.TxtItemId.Name = "TxtItemId";
            this.TxtItemId.ReadOnly = true;
            this.TxtItemId.Size = new System.Drawing.Size(100, 20);
            this.TxtItemId.TabIndex = 0;
            // 
            // TxtListingFee
            // 
            this.TxtListingFee.Location = new System.Drawing.Point(376, 24);
            this.TxtListingFee.Name = "TxtListingFee";
            this.TxtListingFee.ReadOnly = true;
            this.TxtListingFee.Size = new System.Drawing.Size(100, 20);
            this.TxtListingFee.TabIndex = 0;
            // 
            // LblListingFee
            // 
            this.LblListingFee.Location = new System.Drawing.Point(328, 24);
            this.LblListingFee.Name = "LblListingFee";
            this.LblListingFee.Size = new System.Drawing.Size(100, 23);
            this.LblListingFee.TabIndex = 1;
            this.LblListingFee.Text = "Fees:";
            // 
            // TxtReservePrice
            // 
            this.TxtReservePrice.Location = new System.Drawing.Point(272, 160);
            this.TxtReservePrice.Name = "TxtReservePrice";
            this.TxtReservePrice.Size = new System.Drawing.Size(100, 20);
            this.TxtReservePrice.TabIndex = 32;
            this.TxtReservePrice.Text = "2.0";
            // 
            // LblBuyItNowPrice
            // 
            this.LblBuyItNowPrice.Location = new System.Drawing.Point(384, 160);
            this.LblBuyItNowPrice.Name = "LblBuyItNowPrice";
            this.LblBuyItNowPrice.Size = new System.Drawing.Size(100, 23);
            this.LblBuyItNowPrice.TabIndex = 41;
            this.LblBuyItNowPrice.Text = "BIN Price:";
            // 
            // BtnAddItem
            // 
            this.BtnAddItem.Location = new System.Drawing.Point(112, 464);
            this.BtnAddItem.Name = "BtnAddItem";
            this.BtnAddItem.Size = new System.Drawing.Size(75, 23);
            this.BtnAddItem.TabIndex = 36;
            this.BtnAddItem.Text = "AddItem";
            this.BtnAddItem.Click += new System.EventHandler(this.BtnAddItem_Click);
            // 
            // OpenFileDialogIMG
            // 
            this.OpenFileDialogIMG.Filter = "JPEG files (*.jpg)|*.jpg|GIF files (*.gif)|*.gif|All Files|*.*";
            this.OpenFileDialogIMG.FileOk += new System.ComponentModel.CancelEventHandler(this.OpenFileDialogIMG_FileOk);
            // 
            // LblReservePrice
            // 
            this.LblReservePrice.Location = new System.Drawing.Point(184, 160);
            this.LblReservePrice.Name = "LblReservePrice";
            this.LblReservePrice.Size = new System.Drawing.Size(100, 23);
            this.LblReservePrice.TabIndex = 39;
            this.LblReservePrice.Text = "Reserve Price:";
            // 
            // GrpPictures
            // 
            this.GrpPictures.Controls.Add(this.CboPicDisplay);
            this.GrpPictures.Controls.Add(this.BtnAddPic);
            this.GrpPictures.Controls.Add(this.ListPictures);
            this.GrpPictures.Controls.Add(this.BtnRemovePic);
            this.GrpPictures.Location = new System.Drawing.Point(8, 312);
            this.GrpPictures.Name = "GrpPictures";
            this.GrpPictures.Size = new System.Drawing.Size(504, 136);
            this.GrpPictures.TabIndex = 43;
            this.GrpPictures.TabStop = false;
            this.GrpPictures.Text = "Pictures that you want to host in eBay";
            this.GrpPictures.Enter += new System.EventHandler(this.GrpPictures_Enter);
            // 
            // CboPicDisplay
            // 
            this.CboPicDisplay.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CboPicDisplay.Location = new System.Drawing.Point(352, 96);
            this.CboPicDisplay.Name = "CboPicDisplay";
            this.CboPicDisplay.Size = new System.Drawing.Size(121, 21);
            this.CboPicDisplay.TabIndex = 54;
            this.CboPicDisplay.SelectedIndexChanged += new System.EventHandler(this.CboPicDisplay_SelectedIndexChanged);
            // 
            // BtnAddPic
            // 
            this.BtnAddPic.Location = new System.Drawing.Point(360, 24);
            this.BtnAddPic.Name = "BtnAddPic";
            this.BtnAddPic.Size = new System.Drawing.Size(75, 23);
            this.BtnAddPic.TabIndex = 1;
            this.BtnAddPic.Text = "Add...";
            this.BtnAddPic.Click += new System.EventHandler(this.BtnAddPic_Click);
            // 
            // ListPictures
            // 
            this.ListPictures.Location = new System.Drawing.Point(16, 24);
            this.ListPictures.Name = "ListPictures";
            this.ListPictures.Size = new System.Drawing.Size(120, 95);
            this.ListPictures.TabIndex = 0;
            // 
            // BtnRemovePic
            // 
            this.BtnRemovePic.Location = new System.Drawing.Point(360, 56);
            this.BtnRemovePic.Name = "BtnRemovePic";
            this.BtnRemovePic.Size = new System.Drawing.Size(75, 23);
            this.BtnRemovePic.TabIndex = 1;
            this.BtnRemovePic.Text = "Remove";
            this.BtnRemovePic.Click += new System.EventHandler(this.BtnRemovePic_Click);
            // 
            // ChkBoldTitle
            // 
            this.ChkBoldTitle.Location = new System.Drawing.Point(72, 32);
            this.ChkBoldTitle.Name = "ChkBoldTitle";
            this.ChkBoldTitle.Size = new System.Drawing.Size(104, 24);
            this.ChkBoldTitle.TabIndex = 45;
            this.ChkBoldTitle.Text = "Bold";
            // 
            // ChkHighLight
            // 
            this.ChkHighLight.Location = new System.Drawing.Point(128, 32);
            this.ChkHighLight.Name = "ChkHighLight";
            this.ChkHighLight.Size = new System.Drawing.Size(104, 24);
            this.ChkHighLight.TabIndex = 46;
            this.ChkHighLight.Text = "HighLight";
            // 
            // TxtQuantity
            // 
            this.TxtQuantity.Location = new System.Drawing.Point(80, 128);
            this.TxtQuantity.Name = "TxtQuantity";
            this.TxtQuantity.Size = new System.Drawing.Size(100, 20);
            this.TxtQuantity.TabIndex = 47;
            this.TxtQuantity.Text = "1";
            // 
            // LblQuantity
            // 
            this.LblQuantity.Location = new System.Drawing.Point(8, 128);
            this.LblQuantity.Name = "LblQuantity";
            this.LblQuantity.Size = new System.Drawing.Size(100, 23);
            this.LblQuantity.TabIndex = 48;
            this.LblQuantity.Text = "Quantity:";
            // 
            // LblDuration
            // 
            this.LblDuration.Location = new System.Drawing.Point(120, 128);
            this.LblDuration.Name = "LblDuration";
            this.LblDuration.Size = new System.Drawing.Size(100, 23);
            this.LblDuration.TabIndex = 49;
            this.LblDuration.Text = "Duration:";
            // 
            // TxtLocation
            // 
            this.TxtLocation.Location = new System.Drawing.Point(104, 256);
            this.TxtLocation.Name = "TxtLocation";
            this.TxtLocation.Size = new System.Drawing.Size(100, 20);
            this.TxtLocation.TabIndex = 50;
            this.TxtLocation.Text = "San Jose";
            // 
            // LblLocation
            // 
            this.LblLocation.Location = new System.Drawing.Point(8, 256);
            this.LblLocation.Name = "LblLocation";
            this.LblLocation.Size = new System.Drawing.Size(100, 23);
            this.LblLocation.TabIndex = 51;
            this.LblLocation.Text = "Location:";
            // 
            // BtnVerifyAddItem
            // 
            this.BtnVerifyAddItem.Location = new System.Drawing.Point(248, 464);
            this.BtnVerifyAddItem.Name = "BtnVerifyAddItem";
            this.BtnVerifyAddItem.Size = new System.Drawing.Size(75, 23);
            this.BtnVerifyAddItem.TabIndex = 52;
            this.BtnVerifyAddItem.Text = "VerifyAddItem";
            this.BtnVerifyAddItem.Click += new System.EventHandler(this.BtnVerifyAddItem_Click);
            // 
            // CboDuration
            // 
            this.CboDuration.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CboDuration.Location = new System.Drawing.Point(176, 128);
            this.CboDuration.Name = "CboDuration";
            this.CboDuration.Size = new System.Drawing.Size(121, 21);
            this.CboDuration.TabIndex = 53;
            // 
            // ChkEnableBestOffer
            // 
            this.ChkEnableBestOffer.Checked = true;
            this.ChkEnableBestOffer.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ChkEnableBestOffer.Location = new System.Drawing.Point(216, 56);
            this.ChkEnableBestOffer.Name = "ChkEnableBestOffer";
            this.ChkEnableBestOffer.Size = new System.Drawing.Size(104, 24);
            this.ChkEnableBestOffer.TabIndex = 54;
            this.ChkEnableBestOffer.Text = "Enable Best Offer";
            this.ChkEnableBestOffer.Visible = false;
            // 
            // TxtCategory2
            // 
            this.TxtCategory2.Location = new System.Drawing.Point(192, 96);
            this.TxtCategory2.Name = "TxtCategory2";
            this.TxtCategory2.Size = new System.Drawing.Size(100, 20);
            this.TxtCategory2.TabIndex = 57;
            // 
            // lblCategory2
            // 
            this.lblCategory2.Location = new System.Drawing.Point(128, 96);
            this.lblCategory2.Name = "lblCategory2";
            this.lblCategory2.Size = new System.Drawing.Size(100, 23);
            this.lblCategory2.TabIndex = 58;
            this.lblCategory2.Text = "Category2:";
            // 
            // TxtApplicationData
            // 
            this.TxtApplicationData.Location = new System.Drawing.Point(104, 192);
            this.TxtApplicationData.Name = "TxtApplicationData";
            this.TxtApplicationData.Size = new System.Drawing.Size(100, 20);
            this.TxtApplicationData.TabIndex = 59;
            this.TxtApplicationData.Text = "my application data";
            // 
            // lblApplicationData
            // 
            this.lblApplicationData.Location = new System.Drawing.Point(8, 192);
            this.lblApplicationData.Name = "lblApplicationData";
            this.lblApplicationData.Size = new System.Drawing.Size(100, 23);
            this.lblApplicationData.TabIndex = 60;
            this.lblApplicationData.Text = "ApplicationData:";
            // 
            // TxtPayPalEmailAddress
            // 
            this.TxtPayPalEmailAddress.Location = new System.Drawing.Point(330, 221);
            this.TxtPayPalEmailAddress.Name = "TxtPayPalEmailAddress";
            this.TxtPayPalEmailAddress.Size = new System.Drawing.Size(100, 20);
            this.TxtPayPalEmailAddress.TabIndex = 63;
            this.TxtPayPalEmailAddress.Text = "my@test.com";
            // 
            // LblPayPalEmailAddress
            // 
            this.LblPayPalEmailAddress.Location = new System.Drawing.Point(213, 224);
            this.LblPayPalEmailAddress.Name = "LblPayPalEmailAddress";
            this.LblPayPalEmailAddress.Size = new System.Drawing.Size(122, 23);
            this.LblPayPalEmailAddress.TabIndex = 64;
            this.LblPayPalEmailAddress.Text = "PayPal Email Address:";
            // 
            // CboListType
            // 
            this.CboListType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CboListType.Location = new System.Drawing.Point(72, 64);
            this.CboListType.Name = "CboListType";
            this.CboListType.Size = new System.Drawing.Size(121, 21);
            this.CboListType.TabIndex = 66;
            this.CboListType.SelectedIndexChanged += new System.EventHandler(this.CboListType_SelectedIndexChanged);
            // 
            // LblListType
            // 
            this.LblListType.Location = new System.Drawing.Point(8, 64);
            this.LblListType.Name = "LblListType";
            this.LblListType.Size = new System.Drawing.Size(100, 23);
            this.LblListType.TabIndex = 65;
            this.LblListType.Text = "List Type:";
            // 
            // BtnGetCategories
            // 
            this.BtnGetCategories.Location = new System.Drawing.Point(296, 96);
            this.BtnGetCategories.Name = "BtnGetCategories";
            this.BtnGetCategories.Size = new System.Drawing.Size(75, 23);
            this.BtnGetCategories.TabIndex = 67;
            this.BtnGetCategories.Text = "GetCategories ...";
            this.BtnGetCategories.Click += new System.EventHandler(this.BtnGetCategory_Click);
            // 
            // conditionLabel
            // 
            this.conditionLabel.Location = new System.Drawing.Point(213, 257);
            this.conditionLabel.Name = "conditionLabel";
            this.conditionLabel.Size = new System.Drawing.Size(61, 20);
            this.conditionLabel.TabIndex = 64;
            this.conditionLabel.Text = "Condition:";
            // 
            // CboCondition
            // 
            this.CboCondition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CboCondition.Location = new System.Drawing.Point(330, 254);
            this.CboCondition.Name = "CboCondition";
            this.CboCondition.Size = new System.Drawing.Size(101, 21);
            this.CboCondition.TabIndex = 66;
            this.CboCondition.SelectedIndexChanged += new System.EventHandler(this.CboListType_SelectedIndexChanged);
            // 
            // FrmAddItem
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(592, 606);
            this.Controls.Add(this.BtnGetCategories);
            this.Controls.Add(this.CboCondition);
            this.Controls.Add(this.CboListType);
            this.Controls.Add(this.LblListType);
            this.Controls.Add(this.TxtPayPalEmailAddress);
            this.Controls.Add(this.conditionLabel);
            this.Controls.Add(this.LblPayPalEmailAddress);
            this.Controls.Add(this.TxtApplicationData);
            this.Controls.Add(this.lblApplicationData);
            this.Controls.Add(this.TxtCategory2);
            this.Controls.Add(this.lblCategory2);
            this.Controls.Add(this.ChkEnableBestOffer);
            this.Controls.Add(this.CboDuration);
            this.Controls.Add(this.BtnVerifyAddItem);
            this.Controls.Add(this.TxtLocation);
            this.Controls.Add(this.LblLocation);
            this.Controls.Add(this.LblDuration);
            this.Controls.Add(this.TxtQuantity);
            this.Controls.Add(this.LblQuantity);
            this.Controls.Add(this.ChkHighLight);
            this.Controls.Add(this.ChkBoldTitle);
            this.Controls.Add(this.TxtTitle);
            this.Controls.Add(this.TxtDescription);
            this.Controls.Add(this.TxtBuyItNowPrice);
            this.Controls.Add(this.TxtCategory);
            this.Controls.Add(this.TxtReservePrice);
            this.Controls.Add(this.TxtStartPrice);
            this.Controls.Add(this.LblCategory);
            this.Controls.Add(this.LblStartPrice);
            this.Controls.Add(this.LblDescription);
            this.Controls.Add(this.LblTitle);
            this.Controls.Add(this.grpResults);
            this.Controls.Add(this.LblBuyItNowPrice);
            this.Controls.Add(this.BtnAddItem);
            this.Controls.Add(this.LblReservePrice);
            this.Controls.Add(this.GrpPictures);
            this.Name = "FrmAddItem";
            this.Text = "AddItem";
            this.Load += new System.EventHandler(this.FrmAddItem_Load);
            this.grpResults.ResumeLayout(false);
            this.grpResults.PerformLayout();
            this.GrpPictures.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion
		private void FrmAddItem_Load(object sender, System.EventArgs e)
		{
			BtnGetItem.Visible = false;

			string[] enums = {
								 "Days_1",
								 "Days_3",
								 "Days_5",
								 "Days_7"
							 };
			//Enum.GetNames(typeof(ListingDurationCodeType));
			foreach (string item in enums)
			{
				if (item != "CustomCode")
					CboDuration.Items.Add(item);
			}

			enums = Enum.GetNames(typeof(PhotoDisplayCodeType));
			foreach (string item in enums)
			{
				if (item != "CustomCode")
					CboPicDisplay.Items.Add(item);
			}

			enums = Enum.GetNames(typeof(ListingTypeCodeType));
			foreach (string item in enums)
			{
				if (item != "CustomCode" && item != "Unknown")
				{
					CboListType.Items.Add(item);
				}
			}

            CboCondition.Items.Add(new ComboxItem("New", 1000));
            CboCondition.Items.Add(new ComboxItem("Used", 3000));

			CboDuration.SelectedIndex = 0;
			CboPicDisplay.SelectedIndex = 0;
			CboListType.SelectedIndex = 0;
            CboCondition.SelectedIndex = 0;
		}

        public class ComboxItem
        {
            public string Text = "";

            public int Value;
            public ComboxItem(string _Text, int _Value)
            {
                Text = _Text;
                Value = _Value;
            }

            public override string ToString()
            {
                return Text;
            }
        }

		private ItemType FillItem()
		{
			BtnGetItem.Visible = false;

			ItemType item = new ItemType();

			// Set UP Defaults
			item.Currency = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
			item.Country = SiteUtility.GetCountryCodeType(Context.Site);

			item.PaymentMethods = new BuyerPaymentMethodCodeTypeCollection();
			item.PaymentMethods.AddRange(new BuyerPaymentMethodCodeType[] {BuyerPaymentMethodCodeType.PayPal});
			item.RegionID = "0";

			// Set specified values from the form
			item.Title = this.TxtTitle.Text;
			item.Description =  this.TxtDescription.Text;
	
			item.Quantity = Int32.Parse(TxtQuantity.Text, NumberStyles.None);
			item.Location = TxtLocation.Text;

			item.ListingDuration = CboDuration.SelectedItem.ToString();


			item.PrimaryCategory = new CategoryType();
			item.PrimaryCategory.CategoryID = this.TxtCategory.Text;;
			if (TxtStartPrice.Text.Length > 0)
			{
				item.StartPrice = new AmountType();
				item.StartPrice.currencyID = item.Currency;
				item.StartPrice.Value = Convert.ToDouble(this.TxtStartPrice.Text);
				
			}

			if (TxtReservePrice.Visible && TxtReservePrice.Text.Length > 0)
			{
				item.ReservePrice = new AmountType();
				item.ReservePrice.currencyID = item.Currency;
				item.ReservePrice.Value = Convert.ToDouble(this.TxtReservePrice.Text);
			}
			if (TxtBuyItNowPrice.Visible && TxtBuyItNowPrice.Text.Length > 0)
			{
				item.BuyItNowPrice = new AmountType();
				item.BuyItNowPrice.currencyID = item.Currency;
				item.BuyItNowPrice.Value = Convert.ToDouble(this.TxtBuyItNowPrice.Text);
			}

			ListingEnhancementsCodeTypeCollection enhancements = new ListingEnhancementsCodeTypeCollection();
			if (this.ChkBoldTitle.Checked)
				enhancements.Add(ListingEnhancementsCodeType.BoldTitle);
			if (this.ChkHighLight.Checked)
				enhancements.Add(ListingEnhancementsCodeType.Highlight);
				
			if (enhancements.Count > 0)
				item.ListingEnhancement = enhancements;

			item.ListingType = (ListingTypeCodeType)Enum.Parse(typeof(ListingTypeCodeType), CboListType.SelectedItem.ToString());

			if (ChkEnableBestOffer.Visible)
			{
				item.BestOfferDetails = new BestOfferDetailsType();
				item.BestOfferDetails.BestOfferEnabled = ChkEnableBestOffer.Checked;
			}

			if (TxtCategory2.Text.Length > 0)
			{
				item.SecondaryCategory = new CategoryType();
				item.SecondaryCategory.CategoryID = TxtCategory2.Text;
			}

			if (TxtPayPalEmailAddress.Text.Length > 0)
				item.PayPalEmailAddress = TxtPayPalEmailAddress.Text;

			if (TxtApplicationData.Text.Length > 0)
				item.ApplicationData = TxtApplicationData.Text;

            int condition = ((ComboxItem)CboCondition.SelectedItem).Value;
            item.ConditionID = condition;

			//add shipping information
			item.ShippingDetails=getShippingDetails();
			//add handling time
			item.DispatchTimeMax=1;
			//add policy
			item.ReturnPolicy=GetPolicyForUS();
	
			return item;
				
		}

		private void BtnAddItem_Click(object sender, System.EventArgs e)
		{
			try 
			{
				TxtItemId.Text = "";
				TxtListingFee.Text = "";

				ItemType item = FillItem();
				AddItemCall apicall = new AddItemCall(Context);

				if (ListPictures.Items.Count > 0)
				{
					apicall.PictureFileList = new StringCollection();
					item.PictureDetails = new PictureDetailsType();
					item.PictureDetails.PhotoDisplay = (PhotoDisplayCodeType) Enum.Parse(typeof(PhotoDisplayCodeType), CboPicDisplay.SelectedItem.ToString());
				}

				foreach (string pic in ListPictures.Items)
				{
					apicall.PictureFileList.Add(pic);
				}

				FeeTypeCollection fees = apicall.AddItem(item);
	
				TxtItemId.Text = item.ItemID;

				BtnGetItem.Visible = true;
				
				foreach (FeeType fee in fees)
				{
					if (fee.Name == "ListingFee")
					{
						TxtListingFee.Text = fee.Fee.Value.ToString();
						break;
					}
				}

			} 
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void BtnVerifyAddItem_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtItemId.Text = "";
				TxtListingFee.Text = "";
				
				ItemType item = FillItem();
				VerifyAddItemCall apicall = new VerifyAddItemCall(Context);

                // eBay Developer Technical Support Team(DTS) - Devanathan
                // Dated: 24-Feb-2016 9:33PM
                // Following If condition and for loop were commented to add a dummy image URL for bypass upload site hosting call for VerifyAddItem call

                /*
                if (ListPictures.Items.Count > 0)
                {
                    item.PictureDetails = new PictureDetailsType();
                    item.PictureDetails.PhotoDisplay = (PhotoDisplayCodeType)Enum.Parse(typeof(PhotoDisplayCodeType), CboPicDisplay.SelectedItem.ToString());
                }

                foreach (string pic in ListPictures.Items)
                {
                    apicall.PictureFileList.Add(pic);
                }
                */

                // eBay Developer Technical Support Team(DTS) - Devanathan
                // Dated: 24-Feb-2016 9:33PM
                // Added the following four lines of code dummy image URL for bypass upload site hosting call for VerifyAddItem call

                item.PictureDetails = new PictureDetailsType();
                item.PictureDetails.PhotoDisplaySpecified = true;
                item.PictureDetails.PhotoDisplay = PhotoDisplayCodeType.None;
                item.PictureDetails.PictureURL = new StringCollection();
                item.PictureDetails.PictureURL.Add("http://i.ebayimg.com/00/s/MTAwMFgxMDAw/z/Oi4AAOSwzgRWzsz9/$_12.JPG?set_id=880000500F");

				
				FeeTypeCollection fees = apicall.VerifyAddItem(item);
					
				foreach (FeeType fee in fees)
				{
					if (fee.Name == "ListingFee")
					{
						TxtListingFee.Text = fee.Fee.Value.ToString();
						break;
					}
				}
			
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void BtnAddPic_Click(object sender, System.EventArgs e)
		{
				OpenFileDialogIMG.ShowDialog();
		}

		private void BtnRemovePic_Click(object sender, System.EventArgs e)
		{
			if( -1 != ListPictures.SelectedIndex )
				ListPictures.Items.RemoveAt(ListPictures.SelectedIndex);

		}

		private void OpenFileDialogIMG_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
		{
			ListPictures.Items.Add(OpenFileDialogIMG.FileName);

		}

		private void CboListType_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			string selectedText = CboListType.Text;
			bool VisibleFlag = (selectedText.Equals("StoresFixedPrice") || selectedText.Equals("FixedPriceItem")	);
			
			ChkEnableBestOffer.Visible = VisibleFlag;
			TxtReservePrice.Visible = !VisibleFlag;
			TxtBuyItNowPrice.Visible = !VisibleFlag;
			
		}

		private void BtnGetItem_Click(object sender, System.EventArgs e)
		{
			FrmGetItem form = new FrmGetItem();
			form.mItemID = TxtItemId.Text;
			form.Context = Context;
			form.ShowDialog();
		}

		private void BtnGetCategory_Click(object sender, System.EventArgs e)
		{
			FrmGetCategories form = new FrmGetCategories();
			form.Context = Context;
			form.ShowDialog();
		}

		private void GrpPictures_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void CboPicDisplay_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		//set shipping information
		private ShippingDetailsType getShippingDetails()
		{
			// Shipping details.
			ShippingDetailsType sd = new ShippingDetailsType();
			SalesTaxType salesTax = new SalesTaxType();
			salesTax.SalesTaxPercent=0.0825f;
			salesTax.SalesTaxState="CA";
			sd.ApplyShippingDiscount=true;
			AmountType at =new AmountType();
			at.Value=2.8;
			at.currencyID=CurrencyCodeType.USD;			
			sd.PaymentInstructions="eBay DotNet SDK test instruction.";

			// Set calculated shipping.
			sd.ShippingType=ShippingTypeCodeType.Flat;
			//
			ShippingServiceOptionsType st1 = new ShippingServiceOptionsType();
			st1.ShippingService=ShippingServiceCodeType.ShippingMethodStandard.ToString();
			at = new AmountType();
			at.Value=2.0;
			at.currencyID=CurrencyCodeType.USD;
			st1.ShippingServiceAdditionalCost=at;
			at = new AmountType();
			at.Value=1.0;
			at.currencyID=CurrencyCodeType.USD;
			st1.ShippingServiceCost=at;
			st1.ShippingServicePriority=1;
			at = new AmountType();
			at.Value=1.0;
			at.currencyID=CurrencyCodeType.USD;
			st1.ShippingInsuranceCost=at;

			ShippingServiceOptionsType st2 = new ShippingServiceOptionsType();
			st2.ExpeditedService=true;
			st2.ShippingService=ShippingServiceCodeType.ShippingMethodExpress.ToString();
			at = new AmountType();
			at.Value=2.0;
			at.currencyID=CurrencyCodeType.USD;
			st2.ShippingServiceAdditionalCost=at;
			at = new AmountType();
			at.Value=1.0;
			at.currencyID=CurrencyCodeType.USD;
			st2.ShippingServiceCost=at;
			st2.ShippingServicePriority=2;
			at = new AmountType();
			at.Value=1.0;
			at.currencyID=CurrencyCodeType.USD;
			st2.ShippingInsuranceCost=at;

			sd.ShippingServiceOptions=new ShippingServiceOptionsTypeCollection(new ShippingServiceOptionsType[]{st1, st2});

			return sd;
		}
		
		/// <summary>
		/// get a policy for us site only.
		/// </summary>
		/// <returns></returns>
		public static ReturnPolicyType GetPolicyForUS()
		{
			ReturnPolicyType policy=new ReturnPolicyType();
			policy.Refund="MoneyBack";
			policy.ReturnsWithinOption="Days_14";
			policy.ReturnsAcceptedOption="ReturnsAccepted";
			policy.ShippingCostPaidByOption="Buyer";

			return policy;
		}

	}
}
