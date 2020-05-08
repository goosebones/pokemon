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
	public class FrmGetItemTransactions : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.DateTimePicker DatePickModTo;
		private System.Windows.Forms.DateTimePicker DatePickModFrom;
		private System.Windows.Forms.Label LblModSep;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.ColumnHeader ClmPrice;
		private System.Windows.Forms.Button BtnGetItemTransactions;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.Label LblModified;
		private System.Windows.Forms.Label LblTransactions;
		private System.Windows.Forms.ListView LstTransactions;
		private System.Windows.Forms.ColumnHeader ClmTransId;
		private System.Windows.Forms.ColumnHeader ClmAmount;
		private System.Windows.Forms.ColumnHeader ClmQuantity;
		private System.Windows.Forms.ColumnHeader ClmUser;
		private System.Windows.Forms.ColumnHeader ClmBestOfferSale;
		private System.Windows.Forms.Button BtnSendInvoice;
		
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetItemTransactions()
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
			this.ClmTransId = new System.Windows.Forms.ColumnHeader();
			this.ClmPrice = new System.Windows.Forms.ColumnHeader();
			this.ClmAmount = new System.Windows.Forms.ColumnHeader();
			this.ClmQuantity = new System.Windows.Forms.ColumnHeader();
			this.ClmUser = new System.Windows.Forms.ColumnHeader();
			this.ClmBestOfferSale = new System.Windows.Forms.ColumnHeader();
			this.BtnGetItemTransactions = new System.Windows.Forms.Button();
			this.DatePickModTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickModFrom = new System.Windows.Forms.DateTimePicker();
			this.LblModSep = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.LblModified = new System.Windows.Forms.Label();
			this.BtnSendInvoice = new System.Windows.Forms.Button();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblTransactions,
																					this.LstTransactions});
			this.GrpResult.Location = new System.Drawing.Point(8, 120);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(520, 320);
			this.GrpResult.TabIndex = 24;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Results";
			// 
			// LblTransactions
			// 
			this.LblTransactions.Location = new System.Drawing.Point(24, 24);
			this.LblTransactions.Name = "LblTransactions";
			this.LblTransactions.Size = new System.Drawing.Size(112, 23);
			this.LblTransactions.TabIndex = 15;
			this.LblTransactions.Text = "Item Transactions:";
			// 
			// LstTransactions
			// 
			this.LstTransactions.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.LstTransactions.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							  this.ClmTransId,
																							  this.ClmPrice,
																							  this.ClmAmount,
																							  this.ClmQuantity,
																							  this.ClmUser,
																							  this.ClmBestOfferSale});
			this.LstTransactions.GridLines = true;
			this.LstTransactions.Location = new System.Drawing.Point(24, 48);
			this.LstTransactions.Name = "LstTransactions";
			this.LstTransactions.Size = new System.Drawing.Size(472, 256);
			this.LstTransactions.TabIndex = 4;
			this.LstTransactions.View = System.Windows.Forms.View.Details;
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
			this.ClmUser.Width = 80;
			// 
			// ClmBestOfferSale
			// 
			this.ClmBestOfferSale.Text = "BestOfferSale";
			this.ClmBestOfferSale.Width = 80;
			// 
			// BtnGetItemTransactions
			// 
			this.BtnGetItemTransactions.Location = new System.Drawing.Point(152, 80);
			this.BtnGetItemTransactions.Name = "BtnGetItemTransactions";
			this.BtnGetItemTransactions.Size = new System.Drawing.Size(128, 23);
			this.BtnGetItemTransactions.TabIndex = 23;
			this.BtnGetItemTransactions.Text = "GetItemTransactions";
			this.BtnGetItemTransactions.Click += new System.EventHandler(this.BtnGetItemTransactions_Click);
			// 
			// DatePickModTo
			// 
			this.DatePickModTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickModTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickModTo.Location = new System.Drawing.Point(280, 48);
			this.DatePickModTo.Name = "DatePickModTo";
			this.DatePickModTo.Size = new System.Drawing.Size(152, 20);
			this.DatePickModTo.TabIndex = 61;
			// 
			// DatePickModFrom
			// 
			this.DatePickModFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickModFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickModFrom.Location = new System.Drawing.Point(112, 48);
			this.DatePickModFrom.Name = "DatePickModFrom";
			this.DatePickModFrom.Size = new System.Drawing.Size(152, 20);
			this.DatePickModFrom.TabIndex = 60;
			// 
			// LblModSep
			// 
			this.LblModSep.Location = new System.Drawing.Point(264, 48);
			this.LblModSep.Name = "LblModSep";
			this.LblModSep.Size = new System.Drawing.Size(16, 23);
			this.LblModSep.TabIndex = 59;
			this.LblModSep.Text = " - ";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(112, 16);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.TabIndex = 70;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(56, 16);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(56, 16);
			this.LblItemId.TabIndex = 72;
			this.LblItemId.Text = "Item Id:";
			// 
			// LblModified
			// 
			this.LblModified.Location = new System.Drawing.Point(8, 48);
			this.LblModified.Name = "LblModified";
			this.LblModified.Size = new System.Drawing.Size(104, 16);
			this.LblModified.TabIndex = 73;
			this.LblModified.Text = "Modified Between:";
			// 
			// BtnSendInvoice
			// 
			this.BtnSendInvoice.Location = new System.Drawing.Point(320, 80);
			this.BtnSendInvoice.Name = "BtnSendInvoice";
			this.BtnSendInvoice.Size = new System.Drawing.Size(128, 23);
			this.BtnSendInvoice.TabIndex = 74;
			this.BtnSendInvoice.Text = "Send invoice ...";
			this.BtnSendInvoice.Click += new System.EventHandler(this.BtnSendInvoice_Click);
			// 
			// FrmGetItemTransactions
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(544, 453);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.BtnSendInvoice,
																		  this.LblModified,
																		  this.LblItemId,
																		  this.TxtItemId,
																		  this.DatePickModTo,
																		  this.DatePickModFrom,
																		  this.LblModSep,
																		  this.GrpResult,
																		  this.BtnGetItemTransactions});
			this.Name = "FrmGetItemTransactions";
			this.Text = "GetItemTransactions";
			this.Load += new System.EventHandler(this.FrmGetItemTransactions_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void FrmGetItemTransactions_Load(object sender, System.EventArgs e)
		{
			DateTime now = DateTime.Now;
			DatePickModTo.Value = now;
			DatePickModFrom.Value = now.AddDays(-5);

		}
		private void BtnGetItemTransactions_Click(object sender, System.EventArgs e)
		{
			try 
			{
				LstTransactions.Items.Clear();
				GetItemTransactionsCall apicall = new GetItemTransactionsCall(Context);
			
				TransactionTypeCollection transactions = apicall.GetItemTransactions(TxtItemId.Text, DatePickModFrom.Value, DatePickModTo.Value);

				if (transactions.Count == 0)
				{
					MessageBox.Show("There is no transaction");
					return;
				}

				foreach (TransactionType trans in transactions)
				{
					string[] listparams = new string[6];
					listparams[0] = trans.TransactionID;
					listparams[1] = trans.TransactionPrice.Value.ToString();
					listparams[2] = trans.AmountPaid.Value.ToString();
					listparams[3] = trans.QuantityPurchased.ToString();
					listparams[4] = trans.Buyer.UserID;
					listparams[5] = trans.BestOfferSale.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					LstTransactions.Items.Add(vi);

				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void BtnSendInvoice_Click(object sender, System.EventArgs e)
		{
			FrmSendInvoice form = new FrmSendInvoice();
			form.mItemID = TxtItemId.Text;
			form.Context = Context;
			form.ShowDialog();
		}
	}
}
