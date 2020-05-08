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
	public class FrmAddSecondChanceItem : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox grpResults;
		private System.Windows.Forms.TextBox TxtBuyItNowPrice;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblBuyItNowPrice;
		private System.Windows.Forms.Label LblDuration;
		private System.Windows.Forms.ComboBox CboDuration;
		private System.Windows.Forms.TextBox TxtOriginalId;
		private System.Windows.Forms.TextBox TxtRecipient;
		private System.Windows.Forms.Label LblRecipient;
		private System.Windows.Forms.Label LblOriginalId;
		private System.Windows.Forms.Button BtnAddSecondChanceItem;
		private System.Windows.Forms.CheckBox ChkCopyMail;
		private System.Windows.Forms.Button BtnVerifyAddSecondChanceItem;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmAddSecondChanceItem()
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
			this.TxtOriginalId = new System.Windows.Forms.TextBox();
			this.TxtBuyItNowPrice = new System.Windows.Forms.TextBox();
			this.TxtRecipient = new System.Windows.Forms.TextBox();
			this.LblRecipient = new System.Windows.Forms.Label();
			this.LblOriginalId = new System.Windows.Forms.Label();
			this.grpResults = new System.Windows.Forms.GroupBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblBuyItNowPrice = new System.Windows.Forms.Label();
			this.BtnAddSecondChanceItem = new System.Windows.Forms.Button();
			this.ChkCopyMail = new System.Windows.Forms.CheckBox();
			this.LblDuration = new System.Windows.Forms.Label();
			this.BtnVerifyAddSecondChanceItem = new System.Windows.Forms.Button();
			this.CboDuration = new System.Windows.Forms.ComboBox();
			this.grpResults.SuspendLayout();
			this.SuspendLayout();
			// 
			// TxtOriginalId
			// 
			this.TxtOriginalId.Location = new System.Drawing.Point(104, 8);
			this.TxtOriginalId.Name = "TxtOriginalId";
			this.TxtOriginalId.Size = new System.Drawing.Size(80, 20);
			this.TxtOriginalId.TabIndex = 30;
			this.TxtOriginalId.Text = "";
			// 
			// TxtBuyItNowPrice
			// 
			this.TxtBuyItNowPrice.Location = new System.Drawing.Point(104, 72);
			this.TxtBuyItNowPrice.Name = "TxtBuyItNowPrice";
			this.TxtBuyItNowPrice.Size = new System.Drawing.Size(64, 20);
			this.TxtBuyItNowPrice.TabIndex = 34;
			this.TxtBuyItNowPrice.Text = "";
			// 
			// TxtRecipient
			// 
			this.TxtRecipient.Location = new System.Drawing.Point(272, 8);
			this.TxtRecipient.Name = "TxtRecipient";
			this.TxtRecipient.Size = new System.Drawing.Size(72, 20);
			this.TxtRecipient.TabIndex = 31;
			this.TxtRecipient.Text = "";
			// 
			// LblRecipient
			// 
			this.LblRecipient.Location = new System.Drawing.Point(200, 8);
			this.LblRecipient.Name = "LblRecipient";
			this.LblRecipient.Size = new System.Drawing.Size(64, 18);
			this.LblRecipient.TabIndex = 38;
			this.LblRecipient.Text = "Recipient:";
			// 
			// LblOriginalId
			// 
			this.LblOriginalId.Location = new System.Drawing.Point(8, 8);
			this.LblOriginalId.Name = "LblOriginalId";
			this.LblOriginalId.Size = new System.Drawing.Size(88, 18);
			this.LblOriginalId.TabIndex = 37;
			this.LblOriginalId.Text = "Original Item Id:";
			// 
			// grpResults
			// 
			this.grpResults.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.LblItemId,
																					 this.TxtItemId});
			this.grpResults.Location = new System.Drawing.Point(8, 160);
			this.grpResults.Name = "grpResults";
			this.grpResults.Size = new System.Drawing.Size(368, 56);
			this.grpResults.TabIndex = 44;
			this.grpResults.TabStop = false;
			this.grpResults.Text = "Results";
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
			// LblBuyItNowPrice
			// 
			this.LblBuyItNowPrice.Location = new System.Drawing.Point(8, 72);
			this.LblBuyItNowPrice.Name = "LblBuyItNowPrice";
			this.LblBuyItNowPrice.Size = new System.Drawing.Size(64, 18);
			this.LblBuyItNowPrice.TabIndex = 41;
			this.LblBuyItNowPrice.Text = "BIN Price:";
			// 
			// BtnAddSecondChanceItem
			// 
			this.BtnAddSecondChanceItem.Location = new System.Drawing.Point(8, 112);
			this.BtnAddSecondChanceItem.Name = "BtnAddSecondChanceItem";
			this.BtnAddSecondChanceItem.Size = new System.Drawing.Size(176, 26);
			this.BtnAddSecondChanceItem.TabIndex = 36;
			this.BtnAddSecondChanceItem.Text = "AddSecondChanceItem";
			this.BtnAddSecondChanceItem.Click += new System.EventHandler(this.BtnAddSecondChanceItem_Click);
			// 
			// ChkCopyMail
			// 
			this.ChkCopyMail.Location = new System.Drawing.Point(200, 40);
			this.ChkCopyMail.Name = "ChkCopyMail";
			this.ChkCopyMail.Size = new System.Drawing.Size(136, 24);
			this.ChkCopyMail.TabIndex = 45;
			this.ChkCopyMail.Text = "Copy EMail to Seller";
			// 
			// LblDuration
			// 
			this.LblDuration.Location = new System.Drawing.Point(8, 40);
			this.LblDuration.Name = "LblDuration";
			this.LblDuration.Size = new System.Drawing.Size(64, 18);
			this.LblDuration.TabIndex = 49;
			this.LblDuration.Text = "Duration:";
			// 
			// BtnVerifyAddSecondChanceItem
			// 
			this.BtnVerifyAddSecondChanceItem.Location = new System.Drawing.Point(200, 112);
			this.BtnVerifyAddSecondChanceItem.Name = "BtnVerifyAddSecondChanceItem";
			this.BtnVerifyAddSecondChanceItem.Size = new System.Drawing.Size(176, 26);
			this.BtnVerifyAddSecondChanceItem.TabIndex = 52;
			this.BtnVerifyAddSecondChanceItem.Text = "VerifyAddSecondChanceItem";
			this.BtnVerifyAddSecondChanceItem.Click += new System.EventHandler(this.BtnVerifyAddSecondChanceItem_Click);
			// 
			// CboDuration
			// 
			this.CboDuration.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboDuration.Location = new System.Drawing.Point(104, 40);
			this.CboDuration.Name = "CboDuration";
			this.CboDuration.Size = new System.Drawing.Size(64, 21);
			this.CboDuration.TabIndex = 53;
			// 
			// FrmAddSecondChanceItem
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(384, 229);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.CboDuration,
																		  this.BtnVerifyAddSecondChanceItem,
																		  this.LblDuration,
																		  this.ChkCopyMail,
																		  this.TxtOriginalId,
																		  this.TxtBuyItNowPrice,
																		  this.TxtRecipient,
																		  this.LblRecipient,
																		  this.LblOriginalId,
																		  this.grpResults,
																		  this.LblBuyItNowPrice,
																		  this.BtnAddSecondChanceItem});
			this.Name = "FrmAddSecondChanceItem";
			this.Text = "AddSecondChanceItem";
			this.Load += new System.EventHandler(this.FrmAddItem_Load);
			this.grpResults.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
		private void FrmAddItem_Load(object sender, System.EventArgs e)
		{
			string[] enums = {
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
			foreach (string item in enums)
			{
				if (item != "CustomCode")
					CboDuration.Items.Add(item);
			}
			
			CboDuration.SelectedIndex = 0;
		}
		private ItemType FillItem()
		{
			ItemType item = new ItemType();
				
			item.ItemID = TxtOriginalId.Text;
			
			item.ListingDuration = CboDuration.SelectedItem.ToString();


			if (TxtBuyItNowPrice.Text.Length > 0)
			{
				item.BuyItNowPrice = new AmountType();
				item.BuyItNowPrice.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
				item.BuyItNowPrice.Value = Double.Parse(this.TxtBuyItNowPrice.Text, NumberStyles.Currency);
			}

			return item;
		}

		private void BtnAddSecondChanceItem_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtItemId.Text = "";
				
				ItemType item = FillItem();
				AddSecondChanceItemCall apicall = new AddSecondChanceItemCall(Context);

				apicall.RecipientBidderUserID = TxtRecipient.Text;

				apicall.AddSecondChanceItem(item, TxtRecipient.Text);
	
				TxtItemId.Text = item.ItemID;
					
	
				
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void BtnVerifyAddSecondChanceItem_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtItemId.Text = "";
				
				ItemType item = FillItem();
				VerifyAddSecondChanceItemCall apicall = new VerifyAddSecondChanceItemCall(Context);

				apicall.RecipientBidderUserID = TxtRecipient.Text;

				apicall.VerifyAddSecondChanceItem(item, TxtRecipient.Text);
	
				TxtItemId.Text = "";
	

			}			
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
