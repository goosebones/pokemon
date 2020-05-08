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
	public class FrmGetSellerEvents : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnGetSellerEvents;
		private System.Windows.Forms.DateTimePicker DatePickModTo;
		private System.Windows.Forms.DateTimePicker DatePickModFrom;
		private System.Windows.Forms.Label LblModSep;
		private System.Windows.Forms.RadioButton OptModTime;
		private System.Windows.Forms.DateTimePicker DatePickStartTo;
		private System.Windows.Forms.Label LblStartSep;
		private System.Windows.Forms.RadioButton OptStartTime;
		private System.Windows.Forms.DateTimePicker DatePickEndTo;
		private System.Windows.Forms.Label LblEndSep;
		private System.Windows.Forms.RadioButton OptEndTime;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtUserId;
		private System.Windows.Forms.CheckBox ChkIncludeNew;
		private System.Windows.Forms.Label LblUserId;
		private System.Windows.Forms.Label LblEvents;
		private System.Windows.Forms.DateTimePicker DatePickStartFrom;
		private System.Windows.Forms.DateTimePicker DatePickEndFrom;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.ColumnHeader ClmTitle;
		private System.Windows.Forms.ColumnHeader ClmPrice;
		private System.Windows.Forms.ColumnHeader ClmSold;
		private System.Windows.Forms.ColumnHeader ClmBids;
		private System.Windows.Forms.ListView LstEvents;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetSellerEvents()
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
			this.LblEvents = new System.Windows.Forms.Label();
			this.LstEvents = new System.Windows.Forms.ListView();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmTitle = new System.Windows.Forms.ColumnHeader();
			this.ClmPrice = new System.Windows.Forms.ColumnHeader();
			this.ClmSold = new System.Windows.Forms.ColumnHeader();
			this.BtnGetSellerEvents = new System.Windows.Forms.Button();
			this.DatePickModTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickModFrom = new System.Windows.Forms.DateTimePicker();
			this.LblModSep = new System.Windows.Forms.Label();
			this.OptModTime = new System.Windows.Forms.RadioButton();
			this.DatePickStartTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickStartFrom = new System.Windows.Forms.DateTimePicker();
			this.LblStartSep = new System.Windows.Forms.Label();
			this.OptStartTime = new System.Windows.Forms.RadioButton();
			this.DatePickEndTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickEndFrom = new System.Windows.Forms.DateTimePicker();
			this.LblEndSep = new System.Windows.Forms.Label();
			this.OptEndTime = new System.Windows.Forms.RadioButton();
			this.TxtUserId = new System.Windows.Forms.TextBox();
			this.ChkIncludeNew = new System.Windows.Forms.CheckBox();
			this.LblUserId = new System.Windows.Forms.Label();
			this.ClmBids = new System.Windows.Forms.ColumnHeader();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblEvents);
			this.GrpResult.Controls.Add(this.LstEvents);
			this.GrpResult.Location = new System.Drawing.Point(8, 168);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(456, 328);
			this.GrpResult.TabIndex = 24;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Results";
			// 
			// LblEvents
			// 
			this.LblEvents.Location = new System.Drawing.Point(16, 24);
			this.LblEvents.Name = "LblEvents";
			this.LblEvents.Size = new System.Drawing.Size(112, 23);
			this.LblEvents.TabIndex = 15;
			this.LblEvents.Text = "Seller Events:";
			// 
			// LstEvents
			// 
			this.LstEvents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.LstEvents.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.ClmItemId,
																						this.ClmTitle,
																						this.ClmPrice,
																						this.ClmSold,
																						this.ClmBids});
			this.LstEvents.GridLines = true;
			this.LstEvents.Location = new System.Drawing.Point(8, 48);
			this.LstEvents.Name = "LstEvents";
			this.LstEvents.Size = new System.Drawing.Size(440, 264);
			this.LstEvents.TabIndex = 4;
			this.LstEvents.View = System.Windows.Forms.View.Details;
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
			// BtnGetSellerEvents
			// 
			this.BtnGetSellerEvents.Location = new System.Drawing.Point(144, 128);
			this.BtnGetSellerEvents.Name = "BtnGetSellerEvents";
			this.BtnGetSellerEvents.Size = new System.Drawing.Size(128, 23);
			this.BtnGetSellerEvents.TabIndex = 23;
			this.BtnGetSellerEvents.Text = "GetSellerEvents";
			this.BtnGetSellerEvents.Click += new System.EventHandler(this.BtnGetSellerEvents_Click);
			// 
			// DatePickModTo
			// 
			this.DatePickModTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickModTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickModTo.Location = new System.Drawing.Point(312, 48);
			this.DatePickModTo.Name = "DatePickModTo";
			this.DatePickModTo.Size = new System.Drawing.Size(152, 20);
			this.DatePickModTo.TabIndex = 61;
			// 
			// DatePickModFrom
			// 
			this.DatePickModFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickModFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickModFrom.Location = new System.Drawing.Point(144, 48);
			this.DatePickModFrom.Name = "DatePickModFrom";
			this.DatePickModFrom.Size = new System.Drawing.Size(152, 20);
			this.DatePickModFrom.TabIndex = 60;
			// 
			// LblModSep
			// 
			this.LblModSep.Location = new System.Drawing.Point(296, 48);
			this.LblModSep.Name = "LblModSep";
			this.LblModSep.Size = new System.Drawing.Size(16, 23);
			this.LblModSep.TabIndex = 59;
			this.LblModSep.Text = " - ";
			// 
			// OptModTime
			// 
			this.OptModTime.Location = new System.Drawing.Point(8, 48);
			this.OptModTime.Name = "OptModTime";
			this.OptModTime.Size = new System.Drawing.Size(128, 24);
			this.OptModTime.TabIndex = 58;
			this.OptModTime.Text = "Modified Between:";
			this.OptModTime.CheckedChanged += new System.EventHandler(this.Option_CheckedChanged);
			// 
			// DatePickStartTo
			// 
			this.DatePickStartTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickStartTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickStartTo.Location = new System.Drawing.Point(312, 72);
			this.DatePickStartTo.Name = "DatePickStartTo";
			this.DatePickStartTo.Size = new System.Drawing.Size(152, 20);
			this.DatePickStartTo.TabIndex = 65;
			// 
			// DatePickStartFrom
			// 
			this.DatePickStartFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickStartFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickStartFrom.Location = new System.Drawing.Point(144, 72);
			this.DatePickStartFrom.Name = "DatePickStartFrom";
			this.DatePickStartFrom.Size = new System.Drawing.Size(152, 20);
			this.DatePickStartFrom.TabIndex = 64;
			// 
			// LblStartSep
			// 
			this.LblStartSep.Location = new System.Drawing.Point(296, 72);
			this.LblStartSep.Name = "LblStartSep";
			this.LblStartSep.Size = new System.Drawing.Size(16, 23);
			this.LblStartSep.TabIndex = 63;
			this.LblStartSep.Text = " - ";
			// 
			// OptStartTime
			// 
			this.OptStartTime.Location = new System.Drawing.Point(8, 72);
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
			this.DatePickEndTo.Location = new System.Drawing.Point(312, 96);
			this.DatePickEndTo.Name = "DatePickEndTo";
			this.DatePickEndTo.Size = new System.Drawing.Size(152, 20);
			this.DatePickEndTo.TabIndex = 69;
			// 
			// DatePickEndFrom
			// 
			this.DatePickEndFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickEndFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickEndFrom.Location = new System.Drawing.Point(144, 96);
			this.DatePickEndFrom.Name = "DatePickEndFrom";
			this.DatePickEndFrom.Size = new System.Drawing.Size(152, 20);
			this.DatePickEndFrom.TabIndex = 68;
			// 
			// LblEndSep
			// 
			this.LblEndSep.Location = new System.Drawing.Point(296, 96);
			this.LblEndSep.Name = "LblEndSep";
			this.LblEndSep.Size = new System.Drawing.Size(16, 23);
			this.LblEndSep.TabIndex = 67;
			this.LblEndSep.Text = " - ";
			// 
			// OptEndTime
			// 
			this.OptEndTime.Location = new System.Drawing.Point(8, 96);
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
			// ChkIncludeNew
			// 
			this.ChkIncludeNew.Location = new System.Drawing.Point(272, 16);
			this.ChkIncludeNew.Name = "ChkIncludeNew";
			this.ChkIncludeNew.Size = new System.Drawing.Size(128, 24);
			this.ChkIncludeNew.TabIndex = 71;
			this.ChkIncludeNew.Text = "Include New Items";
			// 
			// LblUserId
			// 
			this.LblUserId.Location = new System.Drawing.Point(96, 16);
			this.LblUserId.Name = "LblUserId";
			this.LblUserId.Size = new System.Drawing.Size(56, 16);
			this.LblUserId.TabIndex = 72;
			this.LblUserId.Text = "User Id:";
			// 
			// ClmBids
			// 
			this.ClmBids.Text = "Bids";
			this.ClmBids.Width = 39;
			// 
			// FrmGetSellerEvents
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 501);
			this.Controls.Add(this.LblUserId);
			this.Controls.Add(this.ChkIncludeNew);
			this.Controls.Add(this.TxtUserId);
			this.Controls.Add(this.DatePickEndTo);
			this.Controls.Add(this.DatePickEndFrom);
			this.Controls.Add(this.LblEndSep);
			this.Controls.Add(this.OptEndTime);
			this.Controls.Add(this.DatePickStartTo);
			this.Controls.Add(this.DatePickStartFrom);
			this.Controls.Add(this.LblStartSep);
			this.Controls.Add(this.OptStartTime);
			this.Controls.Add(this.DatePickModTo);
			this.Controls.Add(this.DatePickModFrom);
			this.Controls.Add(this.LblModSep);
			this.Controls.Add(this.OptModTime);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetSellerEvents);
			this.Name = "FrmGetSellerEvents";
			this.Text = "GetSellerEvents";
			this.Load += new System.EventHandler(this.FrmGetSellerList_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

	

		private void FrmGetSellerList_Load(object sender, System.EventArgs e)
		{
			OptModTime.Checked = true;
		
			DateTime now = DateTime.Now;
			DatePickModTo.Value = now;
			DatePickModFrom.Value = now.AddDays(-5);
			DatePickStartTo.Value = now;
			DatePickStartFrom.Value = now.AddDays(-5);
			DatePickEndTo.Value = now.AddDays(5);
			DatePickEndFrom.Value = now;

		}


		private void Option_CheckedChanged(object sender, System.EventArgs e)
		{
			RadioButton sel = (RadioButton) sender;
			if (sel.Checked && sel == OptModTime)
			{
				DatePickModFrom.Enabled = true;
				DatePickModTo.Enabled = true;
				DatePickStartFrom.Enabled = false;
				DatePickStartTo.Enabled = false;
				DatePickEndFrom.Enabled = false;
				DatePickEndTo.Enabled = false;
			} 
			else if (sel.Checked && sel == OptStartTime)
			{
				DatePickModFrom.Enabled = false;
				DatePickModTo.Enabled = false;
				DatePickStartFrom.Enabled = true;
				DatePickStartTo.Enabled = true;
				DatePickEndFrom.Enabled = false;
				DatePickEndTo.Enabled = false;
			}
			else if (sel.Checked && sel == OptEndTime)
			{ 
				DatePickModFrom.Enabled = false;
				DatePickModTo.Enabled = false;
				DatePickStartFrom.Enabled = false;
				DatePickStartTo.Enabled = false;
				DatePickEndFrom.Enabled = true;
				DatePickEndTo.Enabled = true;
			}

		}

		private void BtnGetSellerEvents_Click(object sender, System.EventArgs e)
		{
			try 
			{
				LstEvents.Items.Clear();
				
				GetSellerEventsCall apicall = new GetSellerEventsCall(Context);
			
				if (TxtUserId.Text.Length > 0)
					apicall.UserID = TxtUserId.Text;

				apicall.IncludeNewItem = ChkIncludeNew.Checked;

				if (OptModTime.Checked == true)
				{
					apicall.ModTimeFilter = new TimeFilter(DatePickModFrom.Value, DatePickModTo.Value);
				} 
				else if (OptStartTime.Checked == true)
				{
					apicall.StartTimeFilter = new TimeFilter(DatePickStartFrom.Value, DatePickStartTo.Value);
				}
				else if (OptEndTime.Checked == true)
				{
					apicall.EndTimeFilter = new TimeFilter(DatePickEndFrom.Value, DatePickEndTo.Value);
				}
				apicall.Execute();

				foreach (ItemType evt in apicall.ItemEventList)
				{
					string[] listparams = new string[5];
					listparams[0] = evt.ItemID;
					listparams[1] = evt.Title;
					listparams[2] = evt.SellingStatus.CurrentPrice.Value.ToString();
					listparams[3] = evt.SellingStatus.QuantitySold.ToString();
					listparams[4] = evt.SellingStatus.BidCount.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					LstEvents.Items.Add(vi);

				}


			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
