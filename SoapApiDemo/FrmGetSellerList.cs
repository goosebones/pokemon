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
	public class FrmGetSellerList : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.DateTimePicker DatePickStartTo;
		private System.Windows.Forms.Label LblStartSep;
		private System.Windows.Forms.RadioButton OptStartTime;
		private System.Windows.Forms.DateTimePicker DatePickEndTo;
		private System.Windows.Forms.Label LblEndSep;
		private System.Windows.Forms.RadioButton OptEndTime;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtUserId;
		private System.Windows.Forms.Label LblUserId;
		private System.Windows.Forms.DateTimePicker DatePickStartFrom;
		private System.Windows.Forms.DateTimePicker DatePickEndFrom;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.ColumnHeader ClmTitle;
		private System.Windows.Forms.ColumnHeader ClmPrice;
		private System.Windows.Forms.ColumnHeader ClmSold;
		private System.Windows.Forms.ColumnHeader ClmBids;
		private System.Windows.Forms.ColumnHeader ClmBestOfferEnabled;
		private System.Windows.Forms.Button BtnGetSellerList;
		private System.Windows.Forms.Label LblList;
		private System.Windows.Forms.ListView LstItems;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetSellerList()
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
			this.LblList = new System.Windows.Forms.Label();
			this.LstItems = new System.Windows.Forms.ListView();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmTitle = new System.Windows.Forms.ColumnHeader();
			this.ClmPrice = new System.Windows.Forms.ColumnHeader();
			this.ClmSold = new System.Windows.Forms.ColumnHeader();
			this.ClmBids = new System.Windows.Forms.ColumnHeader();
			this.ClmBestOfferEnabled = new System.Windows.Forms.ColumnHeader();
			this.BtnGetSellerList = new System.Windows.Forms.Button();
			this.DatePickStartTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickStartFrom = new System.Windows.Forms.DateTimePicker();
			this.LblStartSep = new System.Windows.Forms.Label();
			this.OptStartTime = new System.Windows.Forms.RadioButton();
			this.DatePickEndTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickEndFrom = new System.Windows.Forms.DateTimePicker();
			this.LblEndSep = new System.Windows.Forms.Label();
			this.OptEndTime = new System.Windows.Forms.RadioButton();
			this.TxtUserId = new System.Windows.Forms.TextBox();
			this.LblUserId = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblList,
																					this.LstItems});
			this.GrpResult.Location = new System.Drawing.Point(8, 168);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(552, 328);
			this.GrpResult.TabIndex = 24;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Results";
			// 
			// LblList
			// 
			this.LblList.Location = new System.Drawing.Point(16, 24);
			this.LblList.Name = "LblList";
			this.LblList.Size = new System.Drawing.Size(112, 23);
			this.LblList.TabIndex = 15;
			this.LblList.Text = "Seller List:";
			// 
			// LstItems
			// 
			this.LstItems.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.LstItems.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																					   this.ClmItemId,
																					   this.ClmTitle,
																					   this.ClmPrice,
																					   this.ClmSold,
																					   this.ClmBids,
																					   this.ClmBestOfferEnabled});
			this.LstItems.GridLines = true;
			this.LstItems.Location = new System.Drawing.Point(8, 48);
			this.LstItems.Name = "LstItems";
			this.LstItems.Size = new System.Drawing.Size(536, 264);
			this.LstItems.TabIndex = 4;
			this.LstItems.View = System.Windows.Forms.View.Details;
			// 
			// ClmItemId
			// 
			this.ClmItemId.Text = "ItemId";
			this.ClmItemId.Width = 80;
			// 
			// ClmTitle
			// 
			this.ClmTitle.Text = "Title";
			this.ClmTitle.Width = 173;
			// 
			// ClmPrice
			// 
			this.ClmPrice.Text = "Price";
			// 
			// ClmSold
			// 
			this.ClmSold.Text = "Quantity Sold";
			this.ClmSold.Width = 80;
			// 
			// ClmBids
			// 
			this.ClmBids.Text = "Bids";
			this.ClmBids.Width = 39;
			// 
			// ClmBestOfferEnabled
			// 
			this.ClmBestOfferEnabled.Text = "BestOfferEnabled";
			this.ClmBestOfferEnabled.Width = 100;
			// 
			// BtnGetSellerList
			// 
			this.BtnGetSellerList.Location = new System.Drawing.Point(144, 96);
			this.BtnGetSellerList.Name = "BtnGetSellerList";
			this.BtnGetSellerList.Size = new System.Drawing.Size(128, 23);
			this.BtnGetSellerList.TabIndex = 23;
			this.BtnGetSellerList.Text = "GetSellerList";
			this.BtnGetSellerList.Click += new System.EventHandler(this.BtnGetSellerList_Click);
			// 
			// DatePickStartTo
			// 
			this.DatePickStartTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickStartTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickStartTo.Location = new System.Drawing.Point(312, 40);
			this.DatePickStartTo.Name = "DatePickStartTo";
			this.DatePickStartTo.Size = new System.Drawing.Size(152, 20);
			this.DatePickStartTo.TabIndex = 65;
			// 
			// DatePickStartFrom
			// 
			this.DatePickStartFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickStartFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickStartFrom.Location = new System.Drawing.Point(144, 40);
			this.DatePickStartFrom.Name = "DatePickStartFrom";
			this.DatePickStartFrom.Size = new System.Drawing.Size(152, 20);
			this.DatePickStartFrom.TabIndex = 64;
			// 
			// LblStartSep
			// 
			this.LblStartSep.Location = new System.Drawing.Point(296, 48);
			this.LblStartSep.Name = "LblStartSep";
			this.LblStartSep.Size = new System.Drawing.Size(16, 23);
			this.LblStartSep.TabIndex = 63;
			this.LblStartSep.Text = " - ";
			// 
			// OptStartTime
			// 
			this.OptStartTime.Location = new System.Drawing.Point(8, 40);
			this.OptStartTime.Name = "OptStartTime";
			this.OptStartTime.Size = new System.Drawing.Size(128, 24);
			this.OptStartTime.TabIndex = 62;
			this.OptStartTime.Text = "Started Between:";
			this.OptStartTime.CheckedChanged += new System.EventHandler(this.Option_CheckedChanged);
			// 
			// DatePickEndTo
			// 
			this.DatePickEndTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickEndTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickEndTo.Location = new System.Drawing.Point(312, 64);
			this.DatePickEndTo.Name = "DatePickEndTo";
			this.DatePickEndTo.Size = new System.Drawing.Size(152, 20);
			this.DatePickEndTo.TabIndex = 69;
			// 
			// DatePickEndFrom
			// 
			this.DatePickEndFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickEndFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickEndFrom.Location = new System.Drawing.Point(144, 64);
			this.DatePickEndFrom.Name = "DatePickEndFrom";
			this.DatePickEndFrom.Size = new System.Drawing.Size(152, 20);
			this.DatePickEndFrom.TabIndex = 68;
			// 
			// LblEndSep
			// 
			this.LblEndSep.Location = new System.Drawing.Point(296, 72);
			this.LblEndSep.Name = "LblEndSep";
			this.LblEndSep.Size = new System.Drawing.Size(16, 23);
			this.LblEndSep.TabIndex = 67;
			this.LblEndSep.Text = " - ";
			// 
			// OptEndTime
			// 
			this.OptEndTime.Location = new System.Drawing.Point(8, 64);
			this.OptEndTime.Name = "OptEndTime";
			this.OptEndTime.Size = new System.Drawing.Size(128, 24);
			this.OptEndTime.TabIndex = 66;
			this.OptEndTime.Text = "Ended Between:";
			this.OptEndTime.CheckedChanged += new System.EventHandler(this.Option_CheckedChanged);
			// 
			// TxtUserId
			// 
			this.TxtUserId.Location = new System.Drawing.Point(152, 16);
			this.TxtUserId.Name = "TxtUserId";
			this.TxtUserId.TabIndex = 70;
			this.TxtUserId.Text = "";
			// 
			// LblUserId
			// 
			this.LblUserId.Location = new System.Drawing.Point(96, 16);
			this.LblUserId.Name = "LblUserId";
			this.LblUserId.Size = new System.Drawing.Size(56, 16);
			this.LblUserId.TabIndex = 72;
			this.LblUserId.Text = "User Id:";
			// 
			// FrmGetSellerList
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(568, 501);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.LblUserId,
																		  this.TxtUserId,
																		  this.DatePickEndTo,
																		  this.DatePickEndFrom,
																		  this.LblEndSep,
																		  this.OptEndTime,
																		  this.DatePickStartTo,
																		  this.DatePickStartFrom,
																		  this.LblStartSep,
																		  this.OptStartTime,
																		  this.GrpResult,
																		  this.BtnGetSellerList});
			this.Name = "FrmGetSellerList";
			this.Text = "GetSellerList";
			this.Load += new System.EventHandler(this.FrmGetSellerList_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

	

		private void FrmGetSellerList_Load(object sender, System.EventArgs e)
		{
			OptStartTime.Checked = true;
		
			DateTime now = DateTime.Now;
			DatePickStartTo.Value = now;
			DatePickStartFrom.Value = now.AddDays(-5);
			DatePickEndTo.Value = now.AddDays(5);
			DatePickEndFrom.Value = now;

		}


		private void Option_CheckedChanged(object sender, System.EventArgs e)
		{
			RadioButton sel = (RadioButton) sender;
			if (sel.Checked && sel == OptStartTime)
			{
				DatePickStartFrom.Enabled = true;
				DatePickStartTo.Enabled = true;
				DatePickEndFrom.Enabled = false;
				DatePickEndTo.Enabled = false;
			}
			else if (sel.Checked && sel == OptEndTime)
			{ 
				DatePickStartFrom.Enabled = false;
				DatePickStartTo.Enabled = false;
				DatePickEndFrom.Enabled = true;
				DatePickEndTo.Enabled = true;
			}

		}

		private void BtnGetSellerList_Click(object sender, System.EventArgs e)
		{
			try 
			{
				LstItems.Items.Clear();

				GetSellerListCall apicall = new GetSellerListCall(Context);
				apicall.DetailLevelList.Add(DetailLevelCodeType.ReturnAll);
			
				//Pagination is required
				apicall.Pagination = new PaginationType();
				apicall.Pagination.PageNumber = 1;
				apicall.Pagination.EntriesPerPage = 200;

				if (TxtUserId.Text.Length > 0)
					apicall.UserID = TxtUserId.Text;

				if (OptStartTime.Checked == true)
				{
					apicall.StartTimeFilter = new TimeFilter(DatePickStartFrom.Value, DatePickStartTo.Value);
				}
				else if (OptEndTime.Checked == true)
				{
					apicall.EndTimeFilter = new TimeFilter(DatePickEndFrom.Value, DatePickEndTo.Value);
				}

				ItemTypeCollection sellerlist = apicall.GetSellerList();
	
				foreach (ItemType item in sellerlist)
				{
					string[] listparams = new string[6];
					listparams[0] = item.ItemID;
					listparams[1] = item.Title;
					listparams[2] = item.SellingStatus.CurrentPrice.Value.ToString();
					listparams[3] = item.SellingStatus.QuantitySold.ToString();
					listparams[4] = item.SellingStatus.BidCount.ToString();
					
					if (item.BestOfferDetails != null)
						listparams[5] = item.BestOfferDetails.BestOfferEnabled.ToString();
					else
						listparams[5] = "False";


					ListViewItem vi = new ListViewItem(listparams);
					LstItems.Items.Add(vi);

				}


			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}
}
