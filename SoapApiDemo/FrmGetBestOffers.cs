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
	/// Summary description for GetBestOffersForm.
	/// </summary>
	public class FrmGetBestOffers : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Button BtnGetBestOffers;
		private System.Windows.Forms.ListView LstBestOffers;
		private System.Windows.Forms.ColumnHeader ClmBestOfferID;
		private System.Windows.Forms.ColumnHeader ClmExpirationTime;
		private System.Windows.Forms.ColumnHeader ClmPrice;
		private System.Windows.Forms.ColumnHeader ClmStatus;
		private System.Windows.Forms.ColumnHeader ClmQuantity;
		private System.Windows.Forms.ColumnHeader ClmUserID;
		private System.Windows.Forms.ColumnHeader ClmEmail;
		private System.Windows.Forms.ColumnHeader ClmFeedbackScore;
		private System.Windows.Forms.ColumnHeader ClmRegistrationDate;
		private System.Windows.Forms.ColumnHeader ClmUserIDMessage;
		private System.Windows.Forms.ColumnHeader ClmSellerMessage;
		internal System.Windows.Forms.Label LblTitle;
		private System.Windows.Forms.TextBox TxtTitle;
		private System.Windows.Forms.Label LblBestOfferStatus;
		internal System.Windows.Forms.Label LblItemID;
		internal System.Windows.Forms.Label LblBestOfferID;
		private System.Windows.Forms.ComboBox CboFilter;
		private System.Windows.Forms.TextBox TxtBestOfferID;
		private System.Windows.Forms.TextBox TxtItemID;
		internal System.Windows.Forms.Label LblBuyItNowPrice;
		private System.Windows.Forms.TextBox TxtBuyItNowPrice;
		private System.Windows.Forms.TextBox TxtEndTime;
		internal System.Windows.Forms.Label LblEndTime;
		private System.Windows.Forms.TextBox TxtLocation;
		internal System.Windows.Forms.Label LblLocation;
		private System.Windows.Forms.TextBox TxtCurrency;
		internal System.Windows.Forms.Label LblCurrency;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetBestOffers()
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
			this.BtnGetBestOffers = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.TxtCurrency = new System.Windows.Forms.TextBox();
			this.LblCurrency = new System.Windows.Forms.Label();
			this.TxtLocation = new System.Windows.Forms.TextBox();
			this.LblLocation = new System.Windows.Forms.Label();
			this.LblTitle = new System.Windows.Forms.Label();
			this.TxtTitle = new System.Windows.Forms.TextBox();
			this.TxtBuyItNowPrice = new System.Windows.Forms.TextBox();
			this.LblBuyItNowPrice = new System.Windows.Forms.Label();
			this.TxtEndTime = new System.Windows.Forms.TextBox();
			this.LblEndTime = new System.Windows.Forms.Label();
			this.LstBestOffers = new System.Windows.Forms.ListView();
			this.ClmBestOfferID = new System.Windows.Forms.ColumnHeader();
			this.ClmExpirationTime = new System.Windows.Forms.ColumnHeader();
			this.ClmPrice = new System.Windows.Forms.ColumnHeader();
			this.ClmStatus = new System.Windows.Forms.ColumnHeader();
			this.ClmQuantity = new System.Windows.Forms.ColumnHeader();
			this.ClmFeedbackScore = new System.Windows.Forms.ColumnHeader();
			this.ClmRegistrationDate = new System.Windows.Forms.ColumnHeader();
			this.ClmUserID = new System.Windows.Forms.ColumnHeader();
			this.ClmEmail = new System.Windows.Forms.ColumnHeader();
			this.ClmUserIDMessage = new System.Windows.Forms.ColumnHeader();
			this.ClmSellerMessage = new System.Windows.Forms.ColumnHeader();
			this.CboFilter = new System.Windows.Forms.ComboBox();
			this.LblBestOfferStatus = new System.Windows.Forms.Label();
			this.LblItemID = new System.Windows.Forms.Label();
			this.TxtItemID = new System.Windows.Forms.TextBox();
			this.LblBestOfferID = new System.Windows.Forms.Label();
			this.TxtBestOfferID = new System.Windows.Forms.TextBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetBestOffers
			// 
			this.BtnGetBestOffers.Location = new System.Drawing.Point(288, 72);
			this.BtnGetBestOffers.Name = "BtnGetBestOffers";
			this.BtnGetBestOffers.Size = new System.Drawing.Size(176, 23);
			this.BtnGetBestOffers.TabIndex = 9;
			this.BtnGetBestOffers.Text = "GetBestOffers";
			this.BtnGetBestOffers.Click += new System.EventHandler(this.BtnGetBestOffers_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.TxtCurrency,
																					this.LblCurrency,
																					this.TxtLocation,
																					this.LblLocation,
																					this.LblTitle,
																					this.TxtTitle,
																					this.TxtBuyItNowPrice,
																					this.LblBuyItNowPrice,
																					this.TxtEndTime,
																					this.LblEndTime,
																					this.LstBestOffers});
			this.GrpResult.Location = new System.Drawing.Point(8, 104);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(760, 336);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// TxtCurrency
			// 
			this.TxtCurrency.Location = new System.Drawing.Point(552, 64);
			this.TxtCurrency.Name = "TxtCurrency";
			this.TxtCurrency.ReadOnly = true;
			this.TxtCurrency.Size = new System.Drawing.Size(128, 20);
			this.TxtCurrency.TabIndex = 87;
			this.TxtCurrency.Text = "";
			// 
			// LblCurrency
			// 
			this.LblCurrency.Location = new System.Drawing.Point(472, 64);
			this.LblCurrency.Name = "LblCurrency";
			this.LblCurrency.Size = new System.Drawing.Size(64, 23);
			this.LblCurrency.TabIndex = 86;
			this.LblCurrency.Text = "Currency:";
			// 
			// TxtLocation
			// 
			this.TxtLocation.Location = new System.Drawing.Point(552, 24);
			this.TxtLocation.Name = "TxtLocation";
			this.TxtLocation.ReadOnly = true;
			this.TxtLocation.Size = new System.Drawing.Size(128, 20);
			this.TxtLocation.TabIndex = 85;
			this.TxtLocation.Text = "";
			// 
			// LblLocation
			// 
			this.LblLocation.Location = new System.Drawing.Point(472, 24);
			this.LblLocation.Name = "LblLocation";
			this.LblLocation.Size = new System.Drawing.Size(56, 23);
			this.LblLocation.TabIndex = 84;
			this.LblLocation.Text = "Location:";
			// 
			// LblTitle
			// 
			this.LblTitle.Location = new System.Drawing.Point(24, 24);
			this.LblTitle.Name = "LblTitle";
			this.LblTitle.Size = new System.Drawing.Size(72, 23);
			this.LblTitle.TabIndex = 78;
			this.LblTitle.Text = "Title:";
			// 
			// TxtTitle
			// 
			this.TxtTitle.Location = new System.Drawing.Point(96, 24);
			this.TxtTitle.Name = "TxtTitle";
			this.TxtTitle.ReadOnly = true;
			this.TxtTitle.Size = new System.Drawing.Size(312, 20);
			this.TxtTitle.TabIndex = 82;
			this.TxtTitle.Text = "";
			// 
			// TxtBuyItNowPrice
			// 
			this.TxtBuyItNowPrice.Location = new System.Drawing.Point(336, 64);
			this.TxtBuyItNowPrice.Name = "TxtBuyItNowPrice";
			this.TxtBuyItNowPrice.ReadOnly = true;
			this.TxtBuyItNowPrice.Size = new System.Drawing.Size(128, 20);
			this.TxtBuyItNowPrice.TabIndex = 83;
			this.TxtBuyItNowPrice.Text = "";
			// 
			// LblBuyItNowPrice
			// 
			this.LblBuyItNowPrice.Location = new System.Drawing.Point(232, 64);
			this.LblBuyItNowPrice.Name = "LblBuyItNowPrice";
			this.LblBuyItNowPrice.Size = new System.Drawing.Size(96, 23);
			this.LblBuyItNowPrice.TabIndex = 77;
			this.LblBuyItNowPrice.Text = "BuyItNow Price:";
			// 
			// TxtEndTime
			// 
			this.TxtEndTime.Location = new System.Drawing.Point(96, 64);
			this.TxtEndTime.Name = "TxtEndTime";
			this.TxtEndTime.ReadOnly = true;
			this.TxtEndTime.Size = new System.Drawing.Size(120, 20);
			this.TxtEndTime.TabIndex = 81;
			this.TxtEndTime.Text = "";
			// 
			// LblEndTime
			// 
			this.LblEndTime.Location = new System.Drawing.Point(24, 64);
			this.LblEndTime.Name = "LblEndTime";
			this.LblEndTime.Size = new System.Drawing.Size(64, 23);
			this.LblEndTime.TabIndex = 76;
			this.LblEndTime.Text = "End Time:";
			// 
			// LstBestOffers
			// 
			this.LstBestOffers.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							this.ClmBestOfferID,
																							this.ClmExpirationTime,
																							this.ClmPrice,
																							this.ClmStatus,
																							this.ClmQuantity,
																							this.ClmFeedbackScore,
																							this.ClmRegistrationDate,
																							this.ClmUserID,
																							this.ClmEmail,
																							this.ClmUserIDMessage,
																							this.ClmSellerMessage});
			this.LstBestOffers.GridLines = true;
			this.LstBestOffers.Location = new System.Drawing.Point(24, 104);
			this.LstBestOffers.Name = "LstBestOffers";
			this.LstBestOffers.Size = new System.Drawing.Size(720, 216);
			this.LstBestOffers.TabIndex = 13;
			this.LstBestOffers.View = System.Windows.Forms.View.Details;
			// 
			// ClmBestOfferID
			// 
			this.ClmBestOfferID.Text = "Best Offer ID";
			this.ClmBestOfferID.Width = 74;
			// 
			// ClmExpirationTime
			// 
			this.ClmExpirationTime.Text = "Expiration Time";
			this.ClmExpirationTime.Width = 88;
			// 
			// ClmPrice
			// 
			this.ClmPrice.Text = "Price";
			this.ClmPrice.Width = 40;
			// 
			// ClmStatus
			// 
			this.ClmStatus.Text = "Status";
			this.ClmStatus.Width = 44;
			// 
			// ClmQuantity
			// 
			this.ClmQuantity.Text = "Quantity";
			this.ClmQuantity.Width = 56;
			// 
			// ClmFeedbackScore
			// 
			this.ClmFeedbackScore.Text = "Buyer Feedback";
			this.ClmFeedbackScore.Width = 90;
			// 
			// ClmRegistrationDate
			// 
			this.ClmRegistrationDate.Text = "Buyer Reg Date";
			this.ClmRegistrationDate.Width = 90;
			// 
			// ClmUserID
			// 
			this.ClmUserID.Text = "Buyer User ID";
			this.ClmUserID.Width = 80;
			// 
			// ClmEmail
			// 
			this.ClmEmail.Text = "Buyer Email";
			this.ClmEmail.Width = 70;
			// 
			// ClmUserIDMessage
			// 
			this.ClmUserIDMessage.Text = "Buyer Message";
			this.ClmUserIDMessage.Width = 120;
			// 
			// ClmSellerMessage
			// 
			this.ClmSellerMessage.Text = "Seller Message";
			this.ClmSellerMessage.Width = 120;
			// 
			// CboFilter
			// 
			this.CboFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboFilter.Location = new System.Drawing.Point(561, 23);
			this.CboFilter.Name = "CboFilter";
			this.CboFilter.Size = new System.Drawing.Size(144, 21);
			this.CboFilter.TabIndex = 57;
			this.CboFilter.SelectedIndexChanged += new System.EventHandler(this.CboFilter_SelectedIndexChanged);
			// 
			// LblBestOfferStatus
			// 
			this.LblBestOfferStatus.Location = new System.Drawing.Point(434, 24);
			this.LblBestOfferStatus.Name = "LblBestOfferStatus";
			this.LblBestOfferStatus.Size = new System.Drawing.Size(100, 18);
			this.LblBestOfferStatus.TabIndex = 56;
			this.LblBestOfferStatus.Text = "Best Offer Status:";
			// 
			// LblItemID
			// 
			this.LblItemID.Location = new System.Drawing.Point(24, 24);
			this.LblItemID.Name = "LblItemID";
			this.LblItemID.Size = new System.Drawing.Size(48, 18);
			this.LblItemID.TabIndex = 84;
			this.LblItemID.Text = "Item ID:";
			// 
			// TxtItemID
			// 
			this.TxtItemID.Location = new System.Drawing.Point(99, 23);
			this.TxtItemID.Name = "TxtItemID";
			this.TxtItemID.Size = new System.Drawing.Size(88, 20);
			this.TxtItemID.TabIndex = 85;
			this.TxtItemID.Text = "";
			// 
			// LblBestOfferID
			// 
			this.LblBestOfferID.Location = new System.Drawing.Point(214, 24);
			this.LblBestOfferID.Name = "LblBestOfferID";
			this.LblBestOfferID.Size = new System.Drawing.Size(78, 18);
			this.LblBestOfferID.TabIndex = 86;
			this.LblBestOfferID.Text = "Best Offer ID:";
			// 
			// TxtBestOfferID
			// 
			this.TxtBestOfferID.Location = new System.Drawing.Point(319, 23);
			this.TxtBestOfferID.Name = "TxtBestOfferID";
			this.TxtBestOfferID.Size = new System.Drawing.Size(88, 20);
			this.TxtBestOfferID.TabIndex = 87;
			this.TxtBestOfferID.Text = "";
			// 
			// FrmGetBestOffers
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(784, 449);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.LblBestOfferID,
																		  this.TxtBestOfferID,
																		  this.CboFilter,
																		  this.LblBestOfferStatus,
																		  this.GrpResult,
																		  this.BtnGetBestOffers,
																		  this.LblItemID,
																		  this.TxtItemID});
			this.Name = "FrmGetBestOffers";
			this.Text = "GetBestOffersForm";
			this.Load += new System.EventHandler(this.FrmGetBestOffers_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetBestOffers_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstBestOffers.Items.Clear();
	
				GetBestOffersCall apicall = new GetBestOffersCall(Context);

				if (TxtBestOfferID.Text.Length > 0)
					apicall.BestOfferID = TxtBestOfferID.Text;
					
				apicall.BestOfferStatus = (BestOfferStatusCodeType) Enum.Parse(typeof(BestOfferStatusCodeType), CboFilter.SelectedItem.ToString());

				apicall.DetailLevelList.Add(DetailLevelCodeType.ReturnAll);
				string ItemID = TxtItemID.Text;
				BestOfferTypeCollection bestoffers = apicall.GetBestOffers(ItemID);
				
				TxtTitle.Text = apicall.Item.Title;
				TxtLocation.Text = apicall.Item.Location;
				TxtEndTime.Text = apicall.Item.ListingDetails.EndTime.ToString();
				TxtBuyItNowPrice.Text = apicall.Item.BuyItNowPrice.Value.ToString();
				TxtCurrency.Text = apicall.Item.BuyItNowPrice.currencyID.ToString();
				
				foreach (BestOfferType bestoffer in bestoffers)
				{
					string[] listparams = new string[11];
					listparams[0] = bestoffer.BestOfferID;
					listparams[1] = bestoffer.ExpirationTime.ToString();
					listparams[2] =  bestoffer.Price != null ?(bestoffer.Price.Value.ToString()):"";
					listparams[3] =  bestoffer.Status.ToString();
					listparams[4] =  bestoffer.Quantity.ToString();
					listparams[5] =  bestoffer.Buyer.FeedbackScore.ToString();
					listparams[6] =  bestoffer.Buyer.RegistrationDate.ToString();
					listparams[7] =  bestoffer.Buyer.UserID;
					listparams[8] =  bestoffer.Buyer.Email;
					listparams[9] =  bestoffer.BuyerMessage;
					listparams[10] =  bestoffer.SellerMessage;
					
					
					ListViewItem vi = new ListViewItem(listparams);
					this.LstBestOffers.Items.Add(vi);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void FrmGetBestOffers_Load(object sender, System.EventArgs e)
		{
			string[] sorts = Enum.GetNames(typeof(BestOfferStatusCodeType));
			foreach (string srt in sorts)
			{
				if (srt == "All" || srt == "Active")
				{
					CboFilter.Items.Add(srt);
				}
			}
			CboFilter.SelectedIndex = 0;

		}

		private void CboFilter_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

	}
}
