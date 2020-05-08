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
	/// Summary description for GetMyEBayBuyingForm.
	/// </summary>
	public class FrmGetMyeBayBuying : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.ColumnHeader ClmTitle;
		private System.Windows.Forms.ColumnHeader ClmPrice;
		private System.Windows.Forms.ColumnHeader ClmBids;
		private System.Windows.Forms.ColumnHeader ClmTimeLeft;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.ColumnHeader columnHeader2;
		private System.Windows.Forms.ColumnHeader columnHeader3;
		private System.Windows.Forms.ColumnHeader columnHeader4;
		private System.Windows.Forms.ColumnHeader columnHeader5;
		private System.Windows.Forms.ColumnHeader columnHeader6;
		private System.Windows.Forms.ColumnHeader columnHeader7;
		private System.Windows.Forms.ColumnHeader columnHeader8;
		private System.Windows.Forms.ColumnHeader columnHeader9;
		private System.Windows.Forms.ColumnHeader columnHeader10;
		private System.Windows.Forms.ColumnHeader columnHeader11;
		private System.Windows.Forms.ColumnHeader columnHeader12;
		private System.Windows.Forms.ColumnHeader columnHeader13;
		private System.Windows.Forms.ColumnHeader columnHeader14;
		private System.Windows.Forms.ColumnHeader columnHeader15;
		private System.Windows.Forms.Button BtnGetMyeBayBuyingCall;
		private System.Windows.Forms.TextBox TxtMaxItems;
		private System.Windows.Forms.Label LblMaxItems;
		private System.Windows.Forms.Label LblLostItemsList;
		private System.Windows.Forms.ListView LstLostItems;
		private System.Windows.Forms.Label LblWonItems;
		private System.Windows.Forms.ListView LstWonItems;
		private System.Windows.Forms.Label LblBidList;
		private System.Windows.Forms.ListView LstBid;
		private System.Windows.Forms.ListView LstWatchItems;
		private System.Windows.Forms.Label LblWatchList;
		private System.Windows.Forms.ComboBox CboWonSort;
		private System.Windows.Forms.Label LblWonSort;
		private System.Windows.Forms.ComboBox CboBidItemSort;
		private System.Windows.Forms.Label LblBidSort;
		private System.Windows.Forms.ComboBox CboLostSort;
		private System.Windows.Forms.Label LblLostSort;
		private System.Windows.Forms.ComboBox CboWatchSort;
		private System.Windows.Forms.Label LblWatchSort;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetMyeBayBuying()
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
			this.BtnGetMyeBayBuyingCall = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblLostItemsList = new System.Windows.Forms.Label();
			this.LstLostItems = new System.Windows.Forms.ListView();
			this.columnHeader11 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader12 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader13 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader14 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader15 = new System.Windows.Forms.ColumnHeader();
			this.LblWonItems = new System.Windows.Forms.Label();
			this.LstWonItems = new System.Windows.Forms.ListView();
			this.columnHeader6 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader7 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader8 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader9 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader10 = new System.Windows.Forms.ColumnHeader();
			this.LblBidList = new System.Windows.Forms.Label();
			this.LstBid = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader4 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader5 = new System.Windows.Forms.ColumnHeader();
			this.LstWatchItems = new System.Windows.Forms.ListView();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmTitle = new System.Windows.Forms.ColumnHeader();
			this.ClmPrice = new System.Windows.Forms.ColumnHeader();
			this.ClmBids = new System.Windows.Forms.ColumnHeader();
			this.ClmTimeLeft = new System.Windows.Forms.ColumnHeader();
			this.LblWatchList = new System.Windows.Forms.Label();
			this.CboWonSort = new System.Windows.Forms.ComboBox();
			this.LblWonSort = new System.Windows.Forms.Label();
			this.CboBidItemSort = new System.Windows.Forms.ComboBox();
			this.LblBidSort = new System.Windows.Forms.Label();
			this.CboLostSort = new System.Windows.Forms.ComboBox();
			this.LblLostSort = new System.Windows.Forms.Label();
			this.CboWatchSort = new System.Windows.Forms.ComboBox();
			this.LblWatchSort = new System.Windows.Forms.Label();
			this.TxtMaxItems = new System.Windows.Forms.TextBox();
			this.LblMaxItems = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetMyeBayBuyingCall
			// 
			this.BtnGetMyeBayBuyingCall.Location = new System.Drawing.Point(240, 72);
			this.BtnGetMyeBayBuyingCall.Name = "BtnGetMyeBayBuyingCall";
			this.BtnGetMyeBayBuyingCall.Size = new System.Drawing.Size(128, 23);
			this.BtnGetMyeBayBuyingCall.TabIndex = 9;
			this.BtnGetMyeBayBuyingCall.Text = "GetMyeBayBuyingCall";
			this.BtnGetMyeBayBuyingCall.Click += new System.EventHandler(this.BtnGetMyeBayBuyingCall_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblLostItemsList,
																					this.LstLostItems,
																					this.LblWonItems,
																					this.LstWonItems,
																					this.LblBidList,
																					this.LstBid,
																					this.LstWatchItems,
																					this.LblWatchList});
			this.GrpResult.Location = new System.Drawing.Point(8, 104);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(448, 584);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblLostItemsList
			// 
			this.LblLostItemsList.Location = new System.Drawing.Point(16, 432);
			this.LblLostItemsList.Name = "LblLostItemsList";
			this.LblLostItemsList.Size = new System.Drawing.Size(112, 16);
			this.LblLostItemsList.TabIndex = 19;
			this.LblLostItemsList.Text = "Lost Items List:";
			// 
			// LstLostItems
			// 
			this.LstLostItems.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						   this.columnHeader11,
																						   this.columnHeader12,
																						   this.columnHeader13,
																						   this.columnHeader14,
																						   this.columnHeader15});
			this.LstLostItems.GridLines = true;
			this.LstLostItems.Location = new System.Drawing.Point(16, 456);
			this.LstLostItems.Name = "LstLostItems";
			this.LstLostItems.Size = new System.Drawing.Size(416, 112);
			this.LstLostItems.TabIndex = 18;
			this.LstLostItems.View = System.Windows.Forms.View.Details;
			// 
			// columnHeader11
			// 
			this.columnHeader11.Text = "ItemId";
			this.columnHeader11.Width = 81;
			// 
			// columnHeader12
			// 
			this.columnHeader12.Text = "Title";
			this.columnHeader12.Width = 171;
			// 
			// columnHeader13
			// 
			this.columnHeader13.Text = "Price";
			// 
			// columnHeader14
			// 
			this.columnHeader14.Text = "Bids";
			this.columnHeader14.Width = 39;
			// 
			// columnHeader15
			// 
			this.columnHeader15.Text = "EndTime";
			// 
			// LblWonItems
			// 
			this.LblWonItems.Location = new System.Drawing.Point(16, 288);
			this.LblWonItems.Name = "LblWonItems";
			this.LblWonItems.Size = new System.Drawing.Size(112, 16);
			this.LblWonItems.TabIndex = 17;
			this.LblWonItems.Text = "Won Items List:";
			// 
			// LstWonItems
			// 
			this.LstWonItems.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						  this.columnHeader6,
																						  this.columnHeader7,
																						  this.columnHeader8,
																						  this.columnHeader9,
																						  this.columnHeader10});
			this.LstWonItems.GridLines = true;
			this.LstWonItems.Location = new System.Drawing.Point(16, 312);
			this.LstWonItems.Name = "LstWonItems";
			this.LstWonItems.Size = new System.Drawing.Size(416, 112);
			this.LstWonItems.TabIndex = 16;
			this.LstWonItems.View = System.Windows.Forms.View.Details;
			// 
			// columnHeader6
			// 
			this.columnHeader6.Text = "ItemId";
			this.columnHeader6.Width = 81;
			// 
			// columnHeader7
			// 
			this.columnHeader7.Text = "Title";
			this.columnHeader7.Width = 171;
			// 
			// columnHeader8
			// 
			this.columnHeader8.Text = "Price";
			// 
			// columnHeader9
			// 
			this.columnHeader9.Text = "Bids";
			this.columnHeader9.Width = 39;
			// 
			// columnHeader10
			// 
			this.columnHeader10.Text = "EndTime";
			// 
			// LblBidList
			// 
			this.LblBidList.Location = new System.Drawing.Point(16, 16);
			this.LblBidList.Name = "LblBidList";
			this.LblBidList.Size = new System.Drawing.Size(112, 16);
			this.LblBidList.TabIndex = 15;
			this.LblBidList.Text = "Bid Item List:";
			// 
			// LstBid
			// 
			this.LstBid.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																					 this.columnHeader1,
																					 this.columnHeader2,
																					 this.columnHeader3,
																					 this.columnHeader4,
																					 this.columnHeader5});
			this.LstBid.GridLines = true;
			this.LstBid.Location = new System.Drawing.Point(16, 40);
			this.LstBid.Name = "LstBid";
			this.LstBid.Size = new System.Drawing.Size(416, 112);
			this.LstBid.TabIndex = 14;
			this.LstBid.View = System.Windows.Forms.View.Details;
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "ItemId";
			this.columnHeader1.Width = 81;
			// 
			// columnHeader2
			// 
			this.columnHeader2.Text = "Title";
			this.columnHeader2.Width = 171;
			// 
			// columnHeader3
			// 
			this.columnHeader3.Text = "Price";
			// 
			// columnHeader4
			// 
			this.columnHeader4.Text = "Bids";
			this.columnHeader4.Width = 39;
			// 
			// columnHeader5
			// 
			this.columnHeader5.Text = "EndTime";
			// 
			// LstWatchItems
			// 
			this.LstWatchItems.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							this.ClmItemId,
																							this.ClmTitle,
																							this.ClmPrice,
																							this.ClmBids,
																							this.ClmTimeLeft});
			this.LstWatchItems.GridLines = true;
			this.LstWatchItems.Location = new System.Drawing.Point(16, 184);
			this.LstWatchItems.Name = "LstWatchItems";
			this.LstWatchItems.Size = new System.Drawing.Size(416, 96);
			this.LstWatchItems.TabIndex = 13;
			this.LstWatchItems.View = System.Windows.Forms.View.Details;
			// 
			// ClmItemId
			// 
			this.ClmItemId.Text = "ItemId";
			this.ClmItemId.Width = 80;
			// 
			// ClmTitle
			// 
			this.ClmTitle.Text = "Title";
			this.ClmTitle.Width = 171;
			// 
			// ClmPrice
			// 
			this.ClmPrice.Text = "Price";
			// 
			// ClmBids
			// 
			this.ClmBids.Text = "Bids";
			this.ClmBids.Width = 39;
			// 
			// ClmTimeLeft
			// 
			this.ClmTimeLeft.Text = "EndTime";
			// 
			// LblWatchList
			// 
			this.LblWatchList.Location = new System.Drawing.Point(16, 160);
			this.LblWatchList.Name = "LblWatchList";
			this.LblWatchList.Size = new System.Drawing.Size(112, 23);
			this.LblWatchList.TabIndex = 12;
			this.LblWatchList.Text = "Watch Items:";
			// 
			// CboWonSort
			// 
			this.CboWonSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboWonSort.Items.AddRange(new object[] {
															"EndTime",
															"EndTimeDescending"});
			this.CboWonSort.Location = new System.Drawing.Point(112, 32);
			this.CboWonSort.Name = "CboWonSort";
			this.CboWonSort.Size = new System.Drawing.Size(120, 21);
			this.CboWonSort.TabIndex = 55;
			// 
			// LblWonSort
			// 
			this.LblWonSort.Location = new System.Drawing.Point(8, 32);
			this.LblWonSort.Name = "LblWonSort";
			this.LblWonSort.Size = new System.Drawing.Size(104, 18);
			this.LblWonSort.TabIndex = 54;
			this.LblWonSort.Text = "Won Item Sort:";
			// 
			// CboBidItemSort
			// 
			this.CboBidItemSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboBidItemSort.Location = new System.Drawing.Point(112, 8);
			this.CboBidItemSort.Name = "CboBidItemSort";
			this.CboBidItemSort.Size = new System.Drawing.Size(120, 21);
			this.CboBidItemSort.TabIndex = 57;
			// 
			// LblBidSort
			// 
			this.LblBidSort.Location = new System.Drawing.Point(8, 8);
			this.LblBidSort.Name = "LblBidSort";
			this.LblBidSort.Size = new System.Drawing.Size(104, 18);
			this.LblBidSort.TabIndex = 56;
			this.LblBidSort.Text = "Bid Item Sort:";
			// 
			// CboLostSort
			// 
			this.CboLostSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboLostSort.Location = new System.Drawing.Point(344, 32);
			this.CboLostSort.Name = "CboLostSort";
			this.CboLostSort.Size = new System.Drawing.Size(120, 21);
			this.CboLostSort.TabIndex = 61;
			// 
			// LblLostSort
			// 
			this.LblLostSort.Location = new System.Drawing.Point(240, 32);
			this.LblLostSort.Name = "LblLostSort";
			this.LblLostSort.Size = new System.Drawing.Size(104, 18);
			this.LblLostSort.TabIndex = 60;
			this.LblLostSort.Text = "Lost Sort:";
			// 
			// CboWatchSort
			// 
			this.CboWatchSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboWatchSort.Location = new System.Drawing.Point(344, 8);
			this.CboWatchSort.Name = "CboWatchSort";
			this.CboWatchSort.Size = new System.Drawing.Size(120, 21);
			this.CboWatchSort.TabIndex = 59;
			// 
			// LblWatchSort
			// 
			this.LblWatchSort.Location = new System.Drawing.Point(240, 8);
			this.LblWatchSort.Name = "LblWatchSort";
			this.LblWatchSort.Size = new System.Drawing.Size(104, 18);
			this.LblWatchSort.TabIndex = 58;
			this.LblWatchSort.Text = "Watch Item Sort:";
			// 
			// TxtMaxItems
			// 
			this.TxtMaxItems.Location = new System.Drawing.Point(120, 72);
			this.TxtMaxItems.Name = "TxtMaxItems";
			this.TxtMaxItems.Size = new System.Drawing.Size(56, 20);
			this.TxtMaxItems.TabIndex = 62;
			this.TxtMaxItems.Text = "";
			// 
			// LblMaxItems
			// 
			this.LblMaxItems.Location = new System.Drawing.Point(8, 72);
			this.LblMaxItems.Name = "LblMaxItems";
			this.LblMaxItems.Size = new System.Drawing.Size(104, 23);
			this.LblMaxItems.TabIndex = 63;
			this.LblMaxItems.Text = "Max Items Per List:";
			// 
			// FrmGetMyeBayBuying
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 701);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.LblMaxItems,
																		  this.TxtMaxItems,
																		  this.CboBidItemSort,
																		  this.LblBidSort,
																		  this.CboWatchSort,
																		  this.LblWatchSort,
																		  this.CboWonSort,
																		  this.LblWonSort,
																		  this.CboLostSort,
																		  this.LblLostSort,
																		  this.GrpResult,
																		  this.BtnGetMyeBayBuyingCall});
			this.Name = "FrmGetMyeBayBuying";
			this.Text = "GetMyeBayBuyingCall";
			this.Load += new System.EventHandler(this.FrmGetMyeBayBuyingCall_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void FrmGetMyeBayBuyingCall_Load(object sender, System.EventArgs e)
		{
			
			string[] sortnames = Enum.GetNames(typeof(ItemSortTypeCodeType));
			foreach (string sort in sortnames)
			{
				if (sort != "CustomCode")
				{
					CboBidItemSort.Items.Add(sort);
					CboWatchSort.Items.Add(sort);
					CboLostSort.Items.Add(sort);
				}
			}
			CboBidItemSort.SelectedIndex = 0;
			CboWatchSort.SelectedIndex = 0;
			CboWonSort.SelectedIndex = 0;
			CboLostSort.SelectedIndex = 0;
		}

		private void BtnGetMyeBayBuyingCall_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstBid.Items.Clear();
				LstWatchItems.Items.Clear();
				LstWonItems.Items.Clear();
				LstLostItems.Items.Clear();

				GetMyeBayBuyingCall apicall = new GetMyeBayBuyingCall(Context);

				PaginationType pageInfo = null;
				if (TxtMaxItems.Text !=null &&  TxtMaxItems.Text.Length >0) 
				{
					pageInfo = new PaginationType();
					pageInfo.EntriesPerPage = Int32.Parse(TxtMaxItems.Text);
				}

				apicall.BidList = new ItemListCustomizationType();
				apicall.BidList.Sort = (ItemSortTypeCodeType) Enum.Parse(typeof(ItemSortTypeCodeType), CboBidItemSort.SelectedItem.ToString());
				apicall.WatchList = new ItemListCustomizationType();
				apicall.WatchList.Sort = (ItemSortTypeCodeType) Enum.Parse(typeof(ItemSortTypeCodeType), CboWatchSort.SelectedItem.ToString());
				apicall.WonList = new ItemListCustomizationType();
				apicall.WonList.Sort = (ItemSortTypeCodeType) Enum.Parse(typeof(ItemSortTypeCodeType), CboWonSort.SelectedItem.ToString());
				apicall.LostList = new ItemListCustomizationType();
				apicall.LostList.Sort = (ItemSortTypeCodeType) Enum.Parse(typeof(ItemSortTypeCodeType), CboLostSort.SelectedItem.ToString());

				if (pageInfo != null)
				{
					apicall.BidList.Pagination = pageInfo;
					apicall.WatchList.Pagination = pageInfo;
					apicall.WonList.Pagination = pageInfo;
					apicall.LostList.Pagination = pageInfo;
				}

				apicall.GetMyeBayBuying();

			    if (apicall.BidListReturn != null)
				{
					if (apicall.BidListReturn.ItemArray.Count > 0) 
					{
						foreach (ItemType bidItem in apicall.BidListReturn.ItemArray)
						{
							string[] listparams = new string[5];
							listparams[0] = bidItem.ItemID;
							listparams[1] = bidItem.Title;
							if (bidItem.SellingStatus != null ) 
							{
								listparams[2] = bidItem.SellingStatus.CurrentPrice.Value.ToString();
								listparams[3] = bidItem.SellingStatus.BidCount.ToString();
								listparams[4] = bidItem.ListingDetails.EndTime.ToString();
							}
							ListViewItem vi = new ListViewItem(listparams);
							LstBid.Items.Add(vi);

						}
					}
				}
				if (apicall.WatchListReturn != null && 
					apicall.WatchListReturn.ItemArray != null && 
					apicall.WatchListReturn.ItemArray.Count > 0)
				{
					foreach (ItemType watchItem in apicall.WatchListReturn.ItemArray)
					{
						string[] listparams = new string[5];
						listparams[0] = watchItem.ItemID;
						listparams[1] = watchItem.Title;
						if (watchItem.SellingStatus != null ) 
						{
							listparams[2] = watchItem.SellingStatus.CurrentPrice.Value.ToString();
							listparams[3] = watchItem.SellingStatus.BidCount.ToString();
							listparams[4] = watchItem.ListingDetails.EndTime.ToString();
						}
						ListViewItem vi = new ListViewItem(listparams);
						LstWatchItems.Items.Add(vi);

					}
				}
				if (apicall.WonListReturn != null &&
					apicall.WonListReturn.OrderTransactionArray != null &&
					apicall.WonListReturn.OrderTransactionArray.Count > 0)
				{
					foreach (OrderTransactionType wonitem in apicall.WonListReturn.OrderTransactionArray)
					{
						string[] listparams = new string[5];
						if (wonitem.Transaction != null ) 
						{
							listparams[0] = wonitem.Transaction.Item.ItemID;
							listparams[1] = wonitem.Transaction.Item.Title;
							if (wonitem.Transaction.Item.SellingStatus != null ) 
							{
								listparams[2] = wonitem.Transaction.Item.SellingStatus.CurrentPrice.Value.ToString();
								if (wonitem.Transaction.Item.SellingStatus.BidCount >0 )
									listparams[3] = wonitem.Transaction.Item.SellingStatus.BidCount.ToString();
								if (wonitem.Transaction.Item.ListingDetails != null)
									listparams[4] = wonitem.Transaction.Item.ListingDetails.EndTime.ToString();
							}
							ListViewItem vi = new ListViewItem(listparams);
							LstWonItems.Items.Add(vi);

						} 
						else if (wonitem.Order != null) 
						{
							foreach (TransactionType transaction in wonitem.Order.TransactionArray) 
							{
								listparams[0] = transaction.Item.ItemID;
								listparams[1] = transaction.Item.Title;
								if (transaction.Item.SellingStatus != null ) 
								{
									listparams[2] = transaction.Item.SellingStatus.CurrentPrice.Value.ToString();
									if (transaction.Item.SellingStatus.BidCount >0 )
										listparams[3] = transaction.Item.SellingStatus.BidCount.ToString();
									if (transaction.Item.ListingDetails != null)
										listparams[4] = transaction.Item.ListingDetails.EndTime.ToString();
								}
								ListViewItem vi = new ListViewItem(listparams);
								LstWonItems.Items.Add(vi);
							}

						}

					}
				}
				if (apicall.LostListReturn != null &&
					apicall.LostListReturn.ItemArray != null &&
					apicall.LostListReturn.ItemArray.Count > 0)
				{
					foreach (ItemType lostitem in apicall.LostListReturn.ItemArray)
					{
						string[] listparams = new string[5];
						listparams[0] = lostitem.ItemID;
						listparams[1] = lostitem.Title;
						if (lostitem.SellingStatus != null ) 
						{
							listparams[2] = lostitem.SellingStatus.CurrentPrice.Value.ToString();
							listparams[3] = lostitem.SellingStatus.BidCount.ToString();
							listparams[4] = lostitem.ListingDetails.EndTime.ToString();
						}

						ListViewItem vi = new ListViewItem(listparams);
						LstLostItems.Items.Add(vi);

					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}



	
	}
	
}
