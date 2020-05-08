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
	public class FrmRelistItem : System.Windows.Forms.Form
	{
		public ApiContext Context;
		public string mItemID;
		private System.Windows.Forms.GroupBox grpResults;
		private System.Windows.Forms.TextBox TxtStartPrice;
		private System.Windows.Forms.TextBox TxtTitle;
		private System.Windows.Forms.TextBox TxtDescription;
		private System.Windows.Forms.TextBox TxtBuyItNowPrice;
		private System.Windows.Forms.Label LblStartPrice;
		private System.Windows.Forms.Label LblDescription;
		private System.Windows.Forms.Label LblTitle;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.TextBox TxtListingFee;
		private System.Windows.Forms.Label LblListingFee;
		private System.Windows.Forms.TextBox TxtReservePrice;
		private System.Windows.Forms.Label LblBuyItNowPrice;
		private System.Windows.Forms.Label LblReservePrice;
		private System.Windows.Forms.Label LblDuration;
		private System.Windows.Forms.ComboBox CboDuration;
		private System.Windows.Forms.Button BtnRelistItem;
		private System.Windows.Forms.Label LblChangeInstr;
		private System.Windows.Forms.TextBox TxtRelistItemId;
		private System.Windows.Forms.Label LblRelistItemId;
		private System.Windows.Forms.CheckBox ChkApplicationData;
		private System.Windows.Forms.CheckBox ChkCategory2;
		private System.Windows.Forms.CheckBox ChkPayPalEmailAddress;
		private System.Windows.Forms.GroupBox GrpDeleteTag;
		private System.Windows.Forms.Button BtnGetItem;
		private System.Windows.Forms.ComboBox CboEnableBestOffer;
		private System.Windows.Forms.Label LblEnableBestOffer;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ListBox ListPictures;
		private System.Windows.Forms.ComboBox CboPicDisplay;
		private System.Windows.Forms.Button BtnAddPic;
		private System.Windows.Forms.Button BtnRemovePic;
		private System.Windows.Forms.OpenFileDialog OpenFileDialogIMG;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmRelistItem()
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
			this.BtnRelistItem = new System.Windows.Forms.Button();
			this.LblReservePrice = new System.Windows.Forms.Label();
			this.LblDuration = new System.Windows.Forms.Label();
			this.CboDuration = new System.Windows.Forms.ComboBox();
			this.LblChangeInstr = new System.Windows.Forms.Label();
			this.TxtRelistItemId = new System.Windows.Forms.TextBox();
			this.LblRelistItemId = new System.Windows.Forms.Label();
			this.ChkApplicationData = new System.Windows.Forms.CheckBox();
			this.ChkCategory2 = new System.Windows.Forms.CheckBox();
			this.ChkPayPalEmailAddress = new System.Windows.Forms.CheckBox();
			this.GrpDeleteTag = new System.Windows.Forms.GroupBox();
			this.CboEnableBestOffer = new System.Windows.Forms.ComboBox();
			this.LblEnableBestOffer = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.CboPicDisplay = new System.Windows.Forms.ComboBox();
			this.BtnAddPic = new System.Windows.Forms.Button();
			this.BtnRemovePic = new System.Windows.Forms.Button();
			this.ListPictures = new System.Windows.Forms.ListBox();
			this.OpenFileDialogIMG = new System.Windows.Forms.OpenFileDialog();
			this.grpResults.SuspendLayout();
			this.GrpDeleteTag.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// TxtStartPrice
			// 
			this.TxtStartPrice.Location = new System.Drawing.Point(80, 128);
			this.TxtStartPrice.Name = "TxtStartPrice";
			this.TxtStartPrice.Size = new System.Drawing.Size(72, 20);
			this.TxtStartPrice.TabIndex = 33;
			this.TxtStartPrice.Text = "";
			// 
			// TxtTitle
			// 
			this.TxtTitle.Location = new System.Drawing.Point(80, 80);
			this.TxtTitle.Name = "TxtTitle";
			this.TxtTitle.Size = new System.Drawing.Size(392, 20);
			this.TxtTitle.TabIndex = 30;
			this.TxtTitle.Text = "";
			// 
			// TxtDescription
			// 
			this.TxtDescription.Location = new System.Drawing.Point(80, 152);
			this.TxtDescription.Multiline = true;
			this.TxtDescription.Name = "TxtDescription";
			this.TxtDescription.Size = new System.Drawing.Size(392, 56);
			this.TxtDescription.TabIndex = 35;
			this.TxtDescription.Text = "";
			// 
			// TxtBuyItNowPrice
			// 
			this.TxtBuyItNowPrice.Location = new System.Drawing.Point(408, 128);
			this.TxtBuyItNowPrice.Name = "TxtBuyItNowPrice";
			this.TxtBuyItNowPrice.Size = new System.Drawing.Size(64, 20);
			this.TxtBuyItNowPrice.TabIndex = 34;
			this.TxtBuyItNowPrice.Text = "";
			// 
			// LblStartPrice
			// 
			this.LblStartPrice.Location = new System.Drawing.Point(8, 128);
			this.LblStartPrice.Name = "LblStartPrice";
			this.LblStartPrice.Size = new System.Drawing.Size(72, 18);
			this.LblStartPrice.TabIndex = 40;
			this.LblStartPrice.Text = "Start Price:";
			// 
			// LblDescription
			// 
			this.LblDescription.Location = new System.Drawing.Point(8, 152);
			this.LblDescription.Name = "LblDescription";
			this.LblDescription.Size = new System.Drawing.Size(72, 18);
			this.LblDescription.TabIndex = 42;
			this.LblDescription.Text = "Description:";
			// 
			// LblTitle
			// 
			this.LblTitle.Location = new System.Drawing.Point(8, 80);
			this.LblTitle.Name = "LblTitle";
			this.LblTitle.Size = new System.Drawing.Size(72, 18);
			this.LblTitle.TabIndex = 37;
			this.LblTitle.Text = "Title:";
			// 
			// grpResults
			// 
			this.grpResults.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.BtnGetItem,
																					 this.LblItemId,
																					 this.TxtItemId,
																					 this.TxtListingFee,
																					 this.LblListingFee});
			this.grpResults.Location = new System.Drawing.Point(8, 552);
			this.grpResults.Name = "grpResults";
			this.grpResults.Size = new System.Drawing.Size(464, 64);
			this.grpResults.TabIndex = 44;
			this.grpResults.TabStop = false;
			this.grpResults.Text = "Results";
			// 
			// BtnGetItem
			// 
			this.BtnGetItem.Location = new System.Drawing.Point(200, 24);
			this.BtnGetItem.Name = "BtnGetItem";
			this.BtnGetItem.Size = new System.Drawing.Size(80, 26);
			this.BtnGetItem.TabIndex = 37;
			this.BtnGetItem.Text = "GetItem ...";
			this.BtnGetItem.Visible = false;
			this.BtnGetItem.Click += new System.EventHandler(this.BtnGetItem_Click);
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(8, 24);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(48, 23);
			this.LblItemId.TabIndex = 1;
			this.LblItemId.Text = "ItemId:";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(64, 24);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.ReadOnly = true;
			this.TxtItemId.Size = new System.Drawing.Size(120, 20);
			this.TxtItemId.TabIndex = 0;
			this.TxtItemId.Text = "";
			// 
			// TxtListingFee
			// 
			this.TxtListingFee.Location = new System.Drawing.Point(376, 24);
			this.TxtListingFee.Name = "TxtListingFee";
			this.TxtListingFee.ReadOnly = true;
			this.TxtListingFee.Size = new System.Drawing.Size(72, 20);
			this.TxtListingFee.TabIndex = 0;
			this.TxtListingFee.Text = "";
			// 
			// LblListingFee
			// 
			this.LblListingFee.Location = new System.Drawing.Point(328, 24);
			this.LblListingFee.Name = "LblListingFee";
			this.LblListingFee.Size = new System.Drawing.Size(48, 23);
			this.LblListingFee.TabIndex = 1;
			this.LblListingFee.Text = "Fees:";
			// 
			// TxtReservePrice
			// 
			this.TxtReservePrice.Location = new System.Drawing.Point(256, 128);
			this.TxtReservePrice.Name = "TxtReservePrice";
			this.TxtReservePrice.Size = new System.Drawing.Size(64, 20);
			this.TxtReservePrice.TabIndex = 32;
			this.TxtReservePrice.Text = "";
			// 
			// LblBuyItNowPrice
			// 
			this.LblBuyItNowPrice.Location = new System.Drawing.Point(336, 128);
			this.LblBuyItNowPrice.Name = "LblBuyItNowPrice";
			this.LblBuyItNowPrice.Size = new System.Drawing.Size(64, 18);
			this.LblBuyItNowPrice.TabIndex = 41;
			this.LblBuyItNowPrice.Text = "BIN Price:";
			// 
			// BtnRelistItem
			// 
			this.BtnRelistItem.Location = new System.Drawing.Point(176, 512);
			this.BtnRelistItem.Name = "BtnRelistItem";
			this.BtnRelistItem.Size = new System.Drawing.Size(120, 26);
			this.BtnRelistItem.TabIndex = 36;
			this.BtnRelistItem.Text = "RelistItem";
			this.BtnRelistItem.Click += new System.EventHandler(this.BtnRelistItem_Click);
			// 
			// LblReservePrice
			// 
			this.LblReservePrice.Location = new System.Drawing.Point(168, 128);
			this.LblReservePrice.Name = "LblReservePrice";
			this.LblReservePrice.Size = new System.Drawing.Size(80, 18);
			this.LblReservePrice.TabIndex = 39;
			this.LblReservePrice.Text = "Reserve Price:";
			// 
			// LblDuration
			// 
			this.LblDuration.Location = new System.Drawing.Point(8, 104);
			this.LblDuration.Name = "LblDuration";
			this.LblDuration.Size = new System.Drawing.Size(72, 18);
			this.LblDuration.TabIndex = 49;
			this.LblDuration.Text = "Duration:";
			// 
			// CboDuration
			// 
			this.CboDuration.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboDuration.Location = new System.Drawing.Point(80, 104);
			this.CboDuration.Name = "CboDuration";
			this.CboDuration.Size = new System.Drawing.Size(80, 21);
			this.CboDuration.TabIndex = 53;
			// 
			// LblChangeInstr
			// 
			this.LblChangeInstr.Location = new System.Drawing.Point(128, 56);
			this.LblChangeInstr.Name = "LblChangeInstr";
			this.LblChangeInstr.Size = new System.Drawing.Size(280, 18);
			this.LblChangeInstr.TabIndex = 54;
			this.LblChangeInstr.Text = "Change all fields that you want to change for the Item:";
			// 
			// TxtRelistItemId
			// 
			this.TxtRelistItemId.Location = new System.Drawing.Point(240, 16);
			this.TxtRelistItemId.Name = "TxtRelistItemId";
			this.TxtRelistItemId.Size = new System.Drawing.Size(72, 20);
			this.TxtRelistItemId.TabIndex = 55;
			this.TxtRelistItemId.Text = "";
			// 
			// LblRelistItemId
			// 
			this.LblRelistItemId.Location = new System.Drawing.Point(160, 16);
			this.LblRelistItemId.Name = "LblRelistItemId";
			this.LblRelistItemId.Size = new System.Drawing.Size(80, 18);
			this.LblRelistItemId.TabIndex = 56;
			this.LblRelistItemId.Text = "Relist Item Id:";
			// 
			// ChkApplicationData
			// 
			this.ChkApplicationData.Location = new System.Drawing.Point(24, 24);
			this.ChkApplicationData.Name = "ChkApplicationData";
			this.ChkApplicationData.Size = new System.Drawing.Size(160, 24);
			this.ChkApplicationData.TabIndex = 73;
			this.ChkApplicationData.Text = "Application Data";
			// 
			// ChkCategory2
			// 
			this.ChkCategory2.Location = new System.Drawing.Point(24, 48);
			this.ChkCategory2.Name = "ChkCategory2";
			this.ChkCategory2.Size = new System.Drawing.Size(160, 24);
			this.ChkCategory2.TabIndex = 75;
			this.ChkCategory2.Text = "Category2";
			// 
			// ChkPayPalEmailAddress
			// 
			this.ChkPayPalEmailAddress.Location = new System.Drawing.Point(24, 72);
			this.ChkPayPalEmailAddress.Name = "ChkPayPalEmailAddress";
			this.ChkPayPalEmailAddress.Size = new System.Drawing.Size(160, 24);
			this.ChkPayPalEmailAddress.TabIndex = 76;
			this.ChkPayPalEmailAddress.Text = "PayPal Email Address";
			// 
			// GrpDeleteTag
			// 
			this.GrpDeleteTag.Controls.AddRange(new System.Windows.Forms.Control[] {
																					   this.ChkPayPalEmailAddress,
																					   this.ChkCategory2,
																					   this.ChkApplicationData});
			this.GrpDeleteTag.Location = new System.Drawing.Point(16, 256);
			this.GrpDeleteTag.Name = "GrpDeleteTag";
			this.GrpDeleteTag.Size = new System.Drawing.Size(456, 112);
			this.GrpDeleteTag.TabIndex = 80;
			this.GrpDeleteTag.TabStop = false;
			this.GrpDeleteTag.Text = "Delete Selected Data";
			// 
			// CboEnableBestOffer
			// 
			this.CboEnableBestOffer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboEnableBestOffer.Items.AddRange(new object[] {
																	"True",
																	"False"});
			this.CboEnableBestOffer.Location = new System.Drawing.Point(112, 216);
			this.CboEnableBestOffer.Name = "CboEnableBestOffer";
			this.CboEnableBestOffer.Size = new System.Drawing.Size(80, 21);
			this.CboEnableBestOffer.TabIndex = 83;
			// 
			// LblEnableBestOffer
			// 
			this.LblEnableBestOffer.Location = new System.Drawing.Point(8, 216);
			this.LblEnableBestOffer.Name = "LblEnableBestOffer";
			this.LblEnableBestOffer.Size = new System.Drawing.Size(104, 18);
			this.LblEnableBestOffer.TabIndex = 82;
			this.LblEnableBestOffer.Text = "Enable Best Offer:";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.CboPicDisplay,
																					this.BtnAddPic,
																					this.BtnRemovePic,
																					this.ListPictures});
			this.groupBox1.Location = new System.Drawing.Point(16, 376);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(456, 128);
			this.groupBox1.TabIndex = 84;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Pictures that you want to host in eBay";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// CboPicDisplay
			// 
			this.CboPicDisplay.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboPicDisplay.Location = new System.Drawing.Point(344, 96);
			this.CboPicDisplay.Name = "CboPicDisplay";
			this.CboPicDisplay.Size = new System.Drawing.Size(104, 21);
			this.CboPicDisplay.TabIndex = 57;
			this.CboPicDisplay.SelectedIndexChanged += new System.EventHandler(this.CboPicDisplay_SelectedIndexChanged);
			// 
			// BtnAddPic
			// 
			this.BtnAddPic.Location = new System.Drawing.Point(352, 24);
			this.BtnAddPic.Name = "BtnAddPic";
			this.BtnAddPic.Size = new System.Drawing.Size(80, 23);
			this.BtnAddPic.TabIndex = 56;
			this.BtnAddPic.Text = "Add...";
			this.BtnAddPic.Click += new System.EventHandler(this.BtnAddPic_Click);
			// 
			// BtnRemovePic
			// 
			this.BtnRemovePic.Location = new System.Drawing.Point(352, 56);
			this.BtnRemovePic.Name = "BtnRemovePic";
			this.BtnRemovePic.Size = new System.Drawing.Size(80, 23);
			this.BtnRemovePic.TabIndex = 55;
			this.BtnRemovePic.Text = "Remove";
			this.BtnRemovePic.Click += new System.EventHandler(this.BtnRemovePic_Click);
			// 
			// ListPictures
			// 
			this.ListPictures.Location = new System.Drawing.Point(8, 32);
			this.ListPictures.Name = "ListPictures";
			this.ListPictures.Size = new System.Drawing.Size(328, 69);
			this.ListPictures.TabIndex = 1;
			this.ListPictures.SelectedIndexChanged += new System.EventHandler(this.ListPictures_SelectedIndexChanged);
			// 
			// OpenFileDialogIMG
			// 
			this.OpenFileDialogIMG.Filter = "JPEG files (*.jpg)|*.jpg|GIF files (*.gif)|*.gif|All Files|*.*";
			this.OpenFileDialogIMG.FileOk += new System.ComponentModel.CancelEventHandler(this.OpenFileDialogIMG_FileOk);
			// 
			// FrmRelistItem
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(488, 630);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.groupBox1,
																		  this.CboEnableBestOffer,
																		  this.LblEnableBestOffer,
																		  this.GrpDeleteTag,
																		  this.TxtRelistItemId,
																		  this.LblRelistItemId,
																		  this.LblChangeInstr,
																		  this.CboDuration,
																		  this.LblDuration,
																		  this.TxtTitle,
																		  this.TxtDescription,
																		  this.TxtBuyItNowPrice,
																		  this.TxtReservePrice,
																		  this.TxtStartPrice,
																		  this.LblStartPrice,
																		  this.LblDescription,
																		  this.LblTitle,
																		  this.grpResults,
																		  this.LblBuyItNowPrice,
																		  this.BtnRelistItem,
																		  this.LblReservePrice});
			this.Name = "FrmRelistItem";
			this.Text = "RelistItem";
			this.Load += new System.EventHandler(this.FrmRelistItem_Load);
			this.grpResults.ResumeLayout(false);
			this.GrpDeleteTag.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
		private void FrmRelistItem_Load(object sender, System.EventArgs e)
		{
			TxtRelistItemId.Text = mItemID;
			
			string[] durations = {
								 "Days_1",
								 "Days_3",
								 "Days_5",
								 "Days_7",
								 "Days_10",
								 "Days_14",
								 "Days_21",
								 "Days_30",
								 "Days_60",
								 "Days_90",
								 "Days_120",
								 "GTC",
								 "CustomCode"
							 };
			//Enum.GetNames(typeof(ListingDurationCodeType));
			foreach (string day in durations)
			{
				if (day != "CustomCode")
				{
					CboDuration.Items.Add(day);
				}
			}
			CboDuration.SelectedIndex = -1;

			string[] photoEnums = Enum.GetNames(typeof(PhotoDisplayCodeType));
			foreach (string item in photoEnums)
			{
				if (item != "CustomCode")
					CboPicDisplay.Items.Add(item);
			}
			CboPicDisplay.SelectedIndex = 0;
		}


		private void BtnRelistItem_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtItemId.Text = "";
				TxtListingFee.Text = "";
				BtnGetItem.Visible = false;
				
				// Populate the Item
				ItemType item = new ItemType();
			
			
				item.ItemID = TxtRelistItemId.Text;
		

				if (TxtTitle.Text.Length > 0)
				{
					item.Title = this.TxtTitle.Text;
				}
			
				if (TxtDescription.Text.Length > 0)
				{
					item.Description =  this.TxtDescription.Text;
				}

				if (CboDuration.SelectedIndex != -1)
				{
					item.ListingDuration = CboDuration.SelectedItem.ToString();
				}

				CurrencyCodeType currencyCode = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);

				if (TxtStartPrice.Text.Length > 0)
				{
					item.StartPrice = new AmountType();
					item.StartPrice.currencyID = currencyCode;
					item.StartPrice.Value = Double.Parse(this.TxtStartPrice.Text, NumberStyles.Currency);
				}
				if (TxtReservePrice.Text.Length > 0)
				{
					item.ReservePrice = new AmountType();
					item.ReservePrice.currencyID = currencyCode;
					item.ReservePrice.Value = Double.Parse(this.TxtReservePrice.Text, NumberStyles.Currency);
				}
				if (TxtBuyItNowPrice.Text.Length > 0)
				{
					item.BuyItNowPrice = new AmountType();
					item.BuyItNowPrice.currencyID = currencyCode;
					item.BuyItNowPrice.Value = Double.Parse(this.TxtBuyItNowPrice.Text, NumberStyles.Currency);
				}

				if (CboEnableBestOffer.SelectedIndex != -1)
				{
					item.BestOfferDetails = new BestOfferDetailsType();
					item.BestOfferDetails.BestOfferEnabled = Boolean.Parse(CboEnableBestOffer.SelectedItem.ToString());
				}

				StringCollection deletedFields = new StringCollection();
				//drop tag action
				if (ChkCategory2.Checked)
					deletedFields.Add("item.secondaryCategory");

				if (ChkPayPalEmailAddress.Checked)
					deletedFields.Add("item.payPalEmailAddress");

				if (ChkApplicationData.Checked)
					deletedFields.Add("item.applicationData");
		
				
				RelistItemCall apicall = new RelistItemCall(Context);

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

				apicall.DeletedFieldList = deletedFields;

				FeeTypeCollection fees = apicall.RelistItem(item);

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
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void BtnGetItem_Click(object sender, System.EventArgs e)
		{
			FrmGetItem form = new FrmGetItem();
			form.mItemID = TxtItemId.Text;
			form.Context = Context;
			form.ShowDialog();
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void CboPicDisplay_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void BtnAddPic_Click(object sender, System.EventArgs e)
		{
			OpenFileDialogIMG.ShowDialog();
		
		}

		private void OpenFileDialogIMG_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
		{
			ListPictures.Items.Add(OpenFileDialogIMG.FileName);
		}

		private void BtnRemovePic_Click(object sender, System.EventArgs e)
		{
			if( -1 != ListPictures.SelectedIndex )
				ListPictures.Items.RemoveAt(ListPictures.SelectedIndex);		
		}

		private void ListPictures_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		

	}
}
