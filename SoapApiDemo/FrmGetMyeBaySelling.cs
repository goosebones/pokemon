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
	/// Summary description for GetCategoriesForm.
	/// </summary>
	public class FrmGetMyeBaySelling : System.Windows.Forms.Form
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
		private System.Windows.Forms.Button BtnGetMyeBaySellingCall;
		private System.Windows.Forms.Label LblSoldList;
		private System.Windows.Forms.Label LblActiveList;
		private System.Windows.Forms.ListView LstActive;
		private System.Windows.Forms.ListView LstScheduled;
		private System.Windows.Forms.Label LstScheduledList;
		private System.Windows.Forms.ComboBox CboSoldSort;
		private System.Windows.Forms.Label LblSoldSort;
		private System.Windows.Forms.ComboBox CboActiveSort;
		private System.Windows.Forms.Label LblActiveSort;
		private System.Windows.Forms.ComboBox CboUnSoldSort;
		private System.Windows.Forms.Label LblUnSoldSort;
		private System.Windows.Forms.ComboBox CboScheSort;
		private System.Windows.Forms.Label LblScheSort;
		private System.Windows.Forms.ListView LstSold;
		private System.Windows.Forms.Label LblUnSoldList;
		private System.Windows.Forms.ListView LstUnSold;
		private System.Windows.Forms.TextBox TxtMaxItems;
		private System.Windows.Forms.Label LblMaxItems;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetMyeBaySelling()
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
			this.BtnGetMyeBaySellingCall = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblUnSoldList = new System.Windows.Forms.Label();
			this.LstUnSold = new System.Windows.Forms.ListView();
			this.columnHeader11 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader12 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader13 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader14 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader15 = new System.Windows.Forms.ColumnHeader();
			this.LblSoldList = new System.Windows.Forms.Label();
			this.LstSold = new System.Windows.Forms.ListView();
			this.columnHeader6 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader7 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader8 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader9 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader10 = new System.Windows.Forms.ColumnHeader();
			this.LblActiveList = new System.Windows.Forms.Label();
			this.LstActive = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader4 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader5 = new System.Windows.Forms.ColumnHeader();
			this.LstScheduled = new System.Windows.Forms.ListView();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmTitle = new System.Windows.Forms.ColumnHeader();
			this.ClmPrice = new System.Windows.Forms.ColumnHeader();
			this.ClmBids = new System.Windows.Forms.ColumnHeader();
			this.ClmTimeLeft = new System.Windows.Forms.ColumnHeader();
			this.LstScheduledList = new System.Windows.Forms.Label();
			this.CboSoldSort = new System.Windows.Forms.ComboBox();
			this.LblSoldSort = new System.Windows.Forms.Label();
			this.CboActiveSort = new System.Windows.Forms.ComboBox();
			this.LblActiveSort = new System.Windows.Forms.Label();
			this.CboUnSoldSort = new System.Windows.Forms.ComboBox();
			this.LblUnSoldSort = new System.Windows.Forms.Label();
			this.CboScheSort = new System.Windows.Forms.ComboBox();
			this.LblScheSort = new System.Windows.Forms.Label();
			this.TxtMaxItems = new System.Windows.Forms.TextBox();
			this.LblMaxItems = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetMyeBaySellingCall
			// 
			this.BtnGetMyeBaySellingCall.Location = new System.Drawing.Point(240, 72);
			this.BtnGetMyeBaySellingCall.Name = "BtnGetMyeBaySellingCall";
			this.BtnGetMyeBaySellingCall.Size = new System.Drawing.Size(128, 23);
			this.BtnGetMyeBaySellingCall.TabIndex = 9;
			this.BtnGetMyeBaySellingCall.Text = "GetMyeBaySellingCall";
			this.BtnGetMyeBaySellingCall.Click += new System.EventHandler(this.BtnGetMyeBaySellingCall_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblUnSoldList,
																					this.LstUnSold,
																					this.LblSoldList,
																					this.LstSold,
																					this.LblActiveList,
																					this.LstActive,
																					this.LstScheduled,
																					this.LstScheduledList});
			this.GrpResult.Location = new System.Drawing.Point(8, 104);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(448, 584);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblUnSoldList
			// 
			this.LblUnSoldList.Location = new System.Drawing.Point(16, 432);
			this.LblUnSoldList.Name = "LblUnSoldList";
			this.LblUnSoldList.Size = new System.Drawing.Size(112, 16);
			this.LblUnSoldList.TabIndex = 19;
			this.LblUnSoldList.Text = "UnSold List:";
			// 
			// LstUnSold
			// 
			this.LstUnSold.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.columnHeader11,
																						this.columnHeader12,
																						this.columnHeader13,
																						this.columnHeader14,
																						this.columnHeader15});
			this.LstUnSold.GridLines = true;
			this.LstUnSold.Location = new System.Drawing.Point(16, 456);
			this.LstUnSold.Name = "LstUnSold";
			this.LstUnSold.Size = new System.Drawing.Size(416, 112);
			this.LstUnSold.TabIndex = 18;
			this.LstUnSold.View = System.Windows.Forms.View.Details;
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
			// LblSoldList
			// 
			this.LblSoldList.Location = new System.Drawing.Point(16, 288);
			this.LblSoldList.Name = "LblSoldList";
			this.LblSoldList.Size = new System.Drawing.Size(112, 16);
			this.LblSoldList.TabIndex = 17;
			this.LblSoldList.Text = "Sold List:";
			// 
			// LstSold
			// 
			this.LstSold.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																					  this.columnHeader6,
																					  this.columnHeader7,
																					  this.columnHeader8,
																					  this.columnHeader9,
																					  this.columnHeader10});
			this.LstSold.GridLines = true;
			this.LstSold.Location = new System.Drawing.Point(16, 312);
			this.LstSold.Name = "LstSold";
			this.LstSold.Size = new System.Drawing.Size(416, 112);
			this.LstSold.TabIndex = 16;
			this.LstSold.View = System.Windows.Forms.View.Details;
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
			// LblActiveList
			// 
			this.LblActiveList.Location = new System.Drawing.Point(16, 16);
			this.LblActiveList.Name = "LblActiveList";
			this.LblActiveList.Size = new System.Drawing.Size(112, 16);
			this.LblActiveList.TabIndex = 15;
			this.LblActiveList.Text = "Active Item List:";
			// 
			// LstActive
			// 
			this.LstActive.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.columnHeader1,
																						this.columnHeader2,
																						this.columnHeader3,
																						this.columnHeader4,
																						this.columnHeader5});
			this.LstActive.GridLines = true;
			this.LstActive.Location = new System.Drawing.Point(16, 40);
			this.LstActive.Name = "LstActive";
			this.LstActive.Size = new System.Drawing.Size(416, 112);
			this.LstActive.TabIndex = 14;
			this.LstActive.View = System.Windows.Forms.View.Details;
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
			// LstScheduled
			// 
			this.LstScheduled.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						   this.ClmItemId,
																						   this.ClmTitle,
																						   this.ClmPrice,
																						   this.ClmBids,
																						   this.ClmTimeLeft});
			this.LstScheduled.GridLines = true;
			this.LstScheduled.Location = new System.Drawing.Point(16, 184);
			this.LstScheduled.Name = "LstScheduled";
			this.LstScheduled.Size = new System.Drawing.Size(416, 96);
			this.LstScheduled.TabIndex = 13;
			this.LstScheduled.View = System.Windows.Forms.View.Details;
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
			// LstScheduledList
			// 
			this.LstScheduledList.Location = new System.Drawing.Point(16, 160);
			this.LstScheduledList.Name = "LstScheduledList";
			this.LstScheduledList.Size = new System.Drawing.Size(112, 23);
			this.LstScheduledList.TabIndex = 12;
			this.LstScheduledList.Text = "Scheduled Items:";
			// 
			// CboSoldSort
			// 
			this.CboSoldSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboSoldSort.Location = new System.Drawing.Point(112, 32);
			this.CboSoldSort.Name = "CboSoldSort";
			this.CboSoldSort.Size = new System.Drawing.Size(120, 21);
			this.CboSoldSort.TabIndex = 55;
			// 
			// LblSoldSort
			// 
			this.LblSoldSort.Location = new System.Drawing.Point(8, 32);
			this.LblSoldSort.Name = "LblSoldSort";
			this.LblSoldSort.Size = new System.Drawing.Size(104, 18);
			this.LblSoldSort.TabIndex = 54;
			this.LblSoldSort.Text = "Sold Item Sort:";
			// 
			// CboActiveSort
			// 
			this.CboActiveSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboActiveSort.Location = new System.Drawing.Point(112, 8);
			this.CboActiveSort.Name = "CboActiveSort";
			this.CboActiveSort.Size = new System.Drawing.Size(120, 21);
			this.CboActiveSort.TabIndex = 57;
			// 
			// LblActiveSort
			// 
			this.LblActiveSort.Location = new System.Drawing.Point(8, 8);
			this.LblActiveSort.Name = "LblActiveSort";
			this.LblActiveSort.Size = new System.Drawing.Size(104, 18);
			this.LblActiveSort.TabIndex = 56;
			this.LblActiveSort.Text = "Active Item Sort:";
			// 
			// CboUnSoldSort
			// 
			this.CboUnSoldSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboUnSoldSort.Location = new System.Drawing.Point(344, 32);
			this.CboUnSoldSort.Name = "CboUnSoldSort";
			this.CboUnSoldSort.Size = new System.Drawing.Size(120, 21);
			this.CboUnSoldSort.TabIndex = 61;
			// 
			// LblUnSoldSort
			// 
			this.LblUnSoldSort.Location = new System.Drawing.Point(240, 32);
			this.LblUnSoldSort.Name = "LblUnSoldSort";
			this.LblUnSoldSort.Size = new System.Drawing.Size(104, 18);
			this.LblUnSoldSort.TabIndex = 60;
			this.LblUnSoldSort.Text = "UnSold Sort:";
			// 
			// CboScheSort
			// 
			this.CboScheSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboScheSort.Location = new System.Drawing.Point(344, 8);
			this.CboScheSort.Name = "CboScheSort";
			this.CboScheSort.Size = new System.Drawing.Size(120, 21);
			this.CboScheSort.TabIndex = 59;
			// 
			// LblScheSort
			// 
			this.LblScheSort.Location = new System.Drawing.Point(240, 8);
			this.LblScheSort.Name = "LblScheSort";
			this.LblScheSort.Size = new System.Drawing.Size(104, 18);
			this.LblScheSort.TabIndex = 58;
			this.LblScheSort.Text = "Scheduled Sort:";
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
			// FrmGetMyeBaySelling
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 701);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.LblMaxItems,
																		  this.TxtMaxItems,
																		  this.CboUnSoldSort,
																		  this.LblUnSoldSort,
																		  this.CboScheSort,
																		  this.LblScheSort,
																		  this.CboActiveSort,
																		  this.LblActiveSort,
																		  this.CboSoldSort,
																		  this.LblSoldSort,
																		  this.GrpResult,
																		  this.BtnGetMyeBaySellingCall});
			this.Name = "FrmGetMyeBaySelling";
			this.Text = "GetMyeBayCallSelling";
			this.Load += new System.EventHandler(this.FrmGetMyeBaySellingCall_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void FrmGetMyeBaySellingCall_Load(object sender, System.EventArgs e)
		{
			
			string[] sortnames = Enum.GetNames(typeof(ItemSortTypeCodeType));
			foreach (string sort in sortnames)
			{
				if (sort != "CustomCode")
				{
					CboScheSort.Items.Add(sort);
					CboActiveSort.Items.Add(sort);
					CboUnSoldSort.Items.Add(sort);
					CboSoldSort.Items.Add(sort);
				}
			}
			CboScheSort.SelectedIndex = 0;
			CboActiveSort.SelectedIndex = 0;
			CboUnSoldSort.SelectedIndex = 0;
			CboSoldSort.SelectedIndex = 0;
		}

		private void BtnGetMyeBaySellingCall_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstActive.Items.Clear();
				LstScheduled.Items.Clear();
				LstSold.Items.Clear();
				LstUnSold.Items.Clear();

				GetMyeBaySellingCall apicall = new GetMyeBaySellingCall(Context);

				PaginationType pageInfo = null;
				if (TxtMaxItems.Text !=null &&  TxtMaxItems.Text.Length >0) 
				{
					pageInfo = new PaginationType();
					pageInfo.EntriesPerPage = Int32.Parse(TxtMaxItems.Text);
				}

				apicall.ActiveList = new ItemListCustomizationType();
				apicall.ActiveList.Sort = (ItemSortTypeCodeType) Enum.Parse(typeof(ItemSortTypeCodeType), CboActiveSort.SelectedItem.ToString());
				apicall.ScheduledList = new ItemListCustomizationType();
				apicall.ScheduledList.Sort = (ItemSortTypeCodeType) Enum.Parse(typeof(ItemSortTypeCodeType), CboScheSort.SelectedItem.ToString());
				apicall.UnsoldList = new ItemListCustomizationType();
				apicall.UnsoldList.Sort = (ItemSortTypeCodeType) Enum.Parse(typeof(ItemSortTypeCodeType), CboUnSoldSort.SelectedItem.ToString());
				apicall.SoldList = new ItemListCustomizationType();
				apicall.SoldList.Sort = (ItemSortTypeCodeType) Enum.Parse(typeof(ItemSortTypeCodeType), CboSoldSort.SelectedItem.ToString());

				if (pageInfo != null)
				{
					apicall.ActiveList.Pagination = pageInfo;
					apicall.UnsoldList.Pagination = pageInfo;
					apicall.ScheduledList.Pagination = pageInfo;
					apicall.SoldList.Pagination = pageInfo;
				}

				apicall.GetMyeBaySelling();

			    if (apicall.ActiveListReturn != null &&
					apicall.ActiveListReturn.ItemArray != null &&
					apicall.ActiveListReturn.ItemArray.Count > 0)
				{
					foreach (ItemType actitem in apicall.ActiveListReturn.ItemArray)
					{
						string[] listparams = new string[5];
						listparams[0] = actitem.ItemID;
						listparams[1] = actitem.Title;
						if (actitem.SellingStatus != null ) 
						{
							listparams[2] = actitem.SellingStatus.CurrentPrice.Value.ToString();
							listparams[3] = actitem.SellingStatus.BidCount.ToString();
							listparams[4] = actitem.ListingDetails.EndTime.ToString();
						}
						ListViewItem vi = new ListViewItem(listparams);
						LstActive.Items.Add(vi);

					}
				}
				if (apicall.ScheduledListReturn != null &&
					apicall.ScheduledListReturn.ItemArray != null &&
					apicall.ScheduledListReturn.ItemArray.Count > 0)
				{
					foreach (ItemType scheItem in apicall.ScheduledListReturn.ItemArray)
					{
						string[] listparams = new string[5];
						listparams[0] = scheItem.ItemID;
						listparams[1] = scheItem.Title;
						if (scheItem.SellingStatus != null ) 
						{
							listparams[2] = scheItem.SellingStatus.CurrentPrice.Value.ToString();
							listparams[3] = scheItem.SellingStatus.BidCount.ToString();
							listparams[4] = scheItem.ListingDetails.EndTime.ToString();
						}
						ListViewItem vi = new ListViewItem(listparams);
						LstScheduled.Items.Add(vi);

					}
				}
				if (apicall.SoldListReturn != null &&
					apicall.SoldListReturn.OrderTransactionArray != null &&
					apicall.SoldListReturn.OrderTransactionArray.Count > 0)
				{
					foreach (OrderTransactionType solditem in apicall.SoldListReturn.OrderTransactionArray)
					{
						string[] listparams = new string[5];
						if (solditem.Transaction != null ) 
						{
							listparams[0] = solditem.Transaction.Item.ItemID;
							listparams[1] = solditem.Transaction.Item.Title;
							if (solditem.Transaction.Item.SellingStatus != null ) 
							{
								listparams[2] = solditem.Transaction.Item.SellingStatus.CurrentPrice.Value.ToString();
								if (solditem.Transaction.Item.SellingStatus.BidCount >0 )
									listparams[3] = solditem.Transaction.Item.SellingStatus.BidCount.ToString();
								if (solditem.Transaction.Item.ListingDetails != null)
									listparams[4] = solditem.Transaction.Item.ListingDetails.EndTime.ToString();
							}
						}
						ListViewItem vi = new ListViewItem(listparams);
						LstSold.Items.Add(vi);

					}
				}
				if (apicall.UnsoldListReturn != null &&
					apicall.UnsoldListReturn.ItemArray != null &&
					apicall.UnsoldListReturn.ItemArray.Count > 0)
				{
					foreach (ItemType unsolditem in apicall.UnsoldListReturn.ItemArray)
					{
						string[] listparams = new string[5];
						listparams[0] = unsolditem.ItemID;
						listparams[1] = unsolditem.Title;
						if (unsolditem.SellingStatus != null ) 
						{
							listparams[2] = unsolditem.SellingStatus.CurrentPrice.Value.ToString();
							listparams[3] = unsolditem.SellingStatus.BidCount.ToString();
							listparams[4] = unsolditem.ListingDetails.EndTime.ToString();
						}

						ListViewItem vi = new ListViewItem(listparams);
						LstUnSold.Items.Add(vi);

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
