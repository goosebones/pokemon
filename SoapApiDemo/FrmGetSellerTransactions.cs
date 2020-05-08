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
	public class FrmGetSellerTransactions : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.DateTimePicker DatePickModTo;
		private System.Windows.Forms.DateTimePicker DatePickModFrom;
		private System.Windows.Forms.Label LblModSep;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.ColumnHeader ClmPrice;
		private System.Windows.Forms.Label LblModified;
		private System.Windows.Forms.Label LblTransactions;
		private System.Windows.Forms.ListView LstTransactions;
		private System.Windows.Forms.ColumnHeader ClmTransId;
		private System.Windows.Forms.ColumnHeader ClmAmount;
		private System.Windows.Forms.ColumnHeader ClmQuantity;
		private System.Windows.Forms.ColumnHeader ClmUser;
		private System.Windows.Forms.Button BtnGetSellerTransactions;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.CheckBox GMT;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetSellerTransactions()
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
			this.LblTransactions = new System.Windows.Forms.Label();
			this.LstTransactions = new System.Windows.Forms.ListView();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmTransId = new System.Windows.Forms.ColumnHeader();
			this.ClmPrice = new System.Windows.Forms.ColumnHeader();
			this.ClmAmount = new System.Windows.Forms.ColumnHeader();
			this.ClmQuantity = new System.Windows.Forms.ColumnHeader();
			this.ClmUser = new System.Windows.Forms.ColumnHeader();
			this.BtnGetSellerTransactions = new System.Windows.Forms.Button();
			this.DatePickModTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickModFrom = new System.Windows.Forms.DateTimePicker();
			this.LblModSep = new System.Windows.Forms.Label();
			this.LblModified = new System.Windows.Forms.Label();
			this.GMT = new System.Windows.Forms.CheckBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblTransactions,
																					this.LstTransactions});
			this.GrpResult.Location = new System.Drawing.Point(8, 80);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(512, 328);
			this.GrpResult.TabIndex = 24;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Results";
			// 
			// LblTransactions
			// 
			this.LblTransactions.Location = new System.Drawing.Point(16, 24);
			this.LblTransactions.Name = "LblTransactions";
			this.LblTransactions.Size = new System.Drawing.Size(112, 23);
			this.LblTransactions.TabIndex = 15;
			this.LblTransactions.Text = "Seller Transactions:";
			// 
			// LstTransactions
			// 
			this.LstTransactions.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.LstTransactions.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							  this.ClmItemId,
																							  this.ClmTransId,
																							  this.ClmPrice,
																							  this.ClmAmount,
																							  this.ClmQuantity,
																							  this.ClmUser});
			this.LstTransactions.GridLines = true;
			this.LstTransactions.Location = new System.Drawing.Point(8, 48);
			this.LstTransactions.Name = "LstTransactions";
			this.LstTransactions.Size = new System.Drawing.Size(496, 264);
			this.LstTransactions.TabIndex = 4;
			this.LstTransactions.View = System.Windows.Forms.View.Details;
			// 
			// ClmItemId
			// 
			this.ClmItemId.Text = "Item Id";
			this.ClmItemId.Width = 88;
			// 
			// ClmTransId
			// 
			this.ClmTransId.Text = "Transaction Id";
			this.ClmTransId.Width = 100;
			// 
			// ClmPrice
			// 
			this.ClmPrice.Text = "Price";
			// 
			// ClmAmount
			// 
			this.ClmAmount.Text = "Amount";
			this.ClmAmount.Width = 66;
			// 
			// ClmQuantity
			// 
			this.ClmQuantity.Text = "Quantity";
			this.ClmQuantity.Width = 80;
			// 
			// ClmUser
			// 
			this.ClmUser.Text = "User Id";
			this.ClmUser.Width = 94;
			// 
			// BtnGetSellerTransactions
			// 
			this.BtnGetSellerTransactions.Location = new System.Drawing.Point(192, 48);
			this.BtnGetSellerTransactions.Name = "BtnGetSellerTransactions";
			this.BtnGetSellerTransactions.Size = new System.Drawing.Size(128, 23);
			this.BtnGetSellerTransactions.TabIndex = 23;
			this.BtnGetSellerTransactions.Text = "GetSellerTransactions";
			this.BtnGetSellerTransactions.Click += new System.EventHandler(this.BtnGetSellerTransactions_Click);
			// 
			// DatePickModTo
			// 
			this.DatePickModTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickModTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickModTo.Location = new System.Drawing.Point(328, 16);
			this.DatePickModTo.Name = "DatePickModTo";
			this.DatePickModTo.Size = new System.Drawing.Size(152, 20);
			this.DatePickModTo.TabIndex = 61;
			// 
			// DatePickModFrom
			// 
			this.DatePickModFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickModFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickModFrom.Location = new System.Drawing.Point(160, 16);
			this.DatePickModFrom.Name = "DatePickModFrom";
			this.DatePickModFrom.Size = new System.Drawing.Size(152, 20);
			this.DatePickModFrom.TabIndex = 60;
			// 
			// LblModSep
			// 
			this.LblModSep.Location = new System.Drawing.Point(264, 16);
			this.LblModSep.Name = "LblModSep";
			this.LblModSep.Size = new System.Drawing.Size(16, 23);
			this.LblModSep.TabIndex = 59;
			this.LblModSep.Text = " - ";
			// 
			// LblModified
			// 
			this.LblModified.Location = new System.Drawing.Point(56, 16);
			this.LblModified.Name = "LblModified";
			this.LblModified.Size = new System.Drawing.Size(104, 16);
			this.LblModified.TabIndex = 73;
			this.LblModified.Text = "Modified Between:";
			// 
			// GMT
			// 
			this.GMT.Location = new System.Drawing.Point(352, 48);
			this.GMT.Name = "GMT";
			this.GMT.Size = new System.Drawing.Size(72, 24);
			this.GMT.TabIndex = 74;
			this.GMT.Text = "GMT";
			// 
			// FrmGetSellerTransactions
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(528, 413);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.GMT,
																		  this.LblModified,
																		  this.DatePickModTo,
																		  this.DatePickModFrom,
																		  this.LblModSep,
																		  this.GrpResult,
																		  this.BtnGetSellerTransactions});
			this.Name = "FrmGetSellerTransactions";
			this.Text = "GetSellerTransactions";
			this.Load += new System.EventHandler(this.FrmGetSellerTransactions_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void FrmGetSellerTransactions_Load(object sender, System.EventArgs e)
		{
			DateTime now = DateTime.Now;
			DatePickModTo.Value = now;
			DatePickModFrom.Value = now.AddDays(-5);

		}
		private void BtnGetSellerTransactions_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstTransactions.Items.Clear();

				GetSellerTransactionsCall apicall = new GetSellerTransactionsCall(Context);
				TimeFilter timefilter = new TimeFilter();
				if (GMT.Checked.Equals(true))
				{
					timefilter.TimeFromUTC = DatePickModFrom.Value;
					timefilter.TimeToUTC = DatePickModTo.Value;
				} 
				else
				{
					timefilter.TimeFrom = DatePickModFrom.Value;
					timefilter.TimeTo = DatePickModTo.Value;
				}
				TransactionTypeCollection transactions = apicall.GetSellerTransactions(timefilter);
//				TransactionTypeCollection transactions = apicall.GetSellerTransactions(DatePickModFrom.Value, DatePickModTo.Value);
				
				foreach (TransactionType trans in transactions)
				{
					string[] listparams = new string[6];
					listparams[0] = trans.Item.ItemID;
					listparams[1] = trans.TransactionID;
					listparams[2] = trans.TransactionPrice.Value.ToString();
					listparams[3] = trans.AmountPaid.Value.ToString();
					listparams[4] = trans.QuantityPurchased.ToString();
					listparams[5] = trans.Buyer.UserID;

					ListViewItem vi = new ListViewItem(listparams);
					LstTransactions.Items.Add(vi);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}
}
