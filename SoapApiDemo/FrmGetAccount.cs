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
	public class FrmGetAccount : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResults;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button BtnGetAccount;
		private System.Windows.Forms.RadioButton OptLastInvoice;
		private System.Windows.Forms.RadioButton OptInvoiceDate;
		private System.Windows.Forms.RadioButton OptRange;
		private System.Windows.Forms.DateTimePicker DatePickInvoice;
		private System.Windows.Forms.DateTimePicker DatePickFrom;
		private System.Windows.Forms.DateTimePicker DatePickTo;
		private System.Windows.Forms.ListView LstAccountEntries;
		private System.Windows.Forms.ColumnHeader ClmRef;
		private System.Windows.Forms.ColumnHeader ClmType;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.ColumnHeader ClmBalance;
		private System.Windows.Forms.ColumnHeader ClmDate;
		private System.Windows.Forms.Label LblAccState;
		private System.Windows.Forms.TextBox TxtAccState;
		private System.Windows.Forms.TextBox TxtAccBalance;
		private System.Windows.Forms.TextBox TxtPayMethod;
		private System.Windows.Forms.Label LblAccBalance;
		private System.Windows.Forms.Label LblPayMethod;
		private System.Windows.Forms.ColumnHeader ClmNetAmount;
		private System.Windows.Forms.ColumnHeader ClmGrossAmount;
		private System.Windows.Forms.ColumnHeader ClmDescription;
		private System.Windows.Forms.ColumnHeader ClmVAT;
		private System.Windows.Forms.TextBox TxtAmountPastDue;
		private System.Windows.Forms.Label LblAmountPastDue;
		private System.Windows.Forms.Label LblBillingCycleDate;
		private System.Windows.Forms.TextBox TxtBillingCycleDate;
		private System.Windows.Forms.TextBox TxtCreditCardExpiration;
		private System.Windows.Forms.Label LblCreditCardExpiration;
		private System.Windows.Forms.TextBox TxtCreditCardInfo;
		private System.Windows.Forms.Label LblCreditCardInfo;
		private System.Windows.Forms.TextBox TxtCreditCardModifyDate;
		private System.Windows.Forms.Label LblCreditCardModifyDate;
		private System.Windows.Forms.Label LblLastAmountPaid;
		private System.Windows.Forms.TextBox TxtLastAmountPaid;
		private System.Windows.Forms.TextBox TxtPastDue;
		private System.Windows.Forms.Label LblPastDue;
		private System.Windows.Forms.TextBox TxtAccountID;
		private System.Windows.Forms.Label LblAccountID;
		private System.Windows.Forms.GroupBox GrpAccountEntry;
		private System.Windows.Forms.TextBox TxtEntriesPerPage;
		private System.Windows.Forms.Label LblEntryCount;
		private System.Windows.Forms.Label LblPageNumber;
		private System.Windows.Forms.TextBox TxtPageNumber;
		private System.Windows.Forms.Label LblTotalPages;
		private System.Windows.Forms.TextBox TxtTotalNumberOfPages;
		private System.Windows.Forms.TextBox TxtLastPaymentDate;
		private System.Windows.Forms.Label LblLastPaymentDate;
		private System.Windows.Forms.Button BtnNextPage;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetAccount()
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
			this.GrpResults = new System.Windows.Forms.GroupBox();
			this.GrpAccountEntry = new System.Windows.Forms.GroupBox();
			this.LstAccountEntries = new System.Windows.Forms.ListView();
			this.ClmRef = new System.Windows.Forms.ColumnHeader();
			this.ClmType = new System.Windows.Forms.ColumnHeader();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmDescription = new System.Windows.Forms.ColumnHeader();
			this.ClmNetAmount = new System.Windows.Forms.ColumnHeader();
			this.ClmGrossAmount = new System.Windows.Forms.ColumnHeader();
			this.ClmVAT = new System.Windows.Forms.ColumnHeader();
			this.ClmBalance = new System.Windows.Forms.ColumnHeader();
			this.ClmDate = new System.Windows.Forms.ColumnHeader();
			this.TxtAccountID = new System.Windows.Forms.TextBox();
			this.LblAccountID = new System.Windows.Forms.Label();
			this.TxtPastDue = new System.Windows.Forms.TextBox();
			this.LblPastDue = new System.Windows.Forms.Label();
			this.LblLastAmountPaid = new System.Windows.Forms.Label();
			this.TxtLastAmountPaid = new System.Windows.Forms.TextBox();
			this.TxtCreditCardModifyDate = new System.Windows.Forms.TextBox();
			this.LblCreditCardModifyDate = new System.Windows.Forms.Label();
			this.TxtCreditCardInfo = new System.Windows.Forms.TextBox();
			this.LblCreditCardInfo = new System.Windows.Forms.Label();
			this.TxtCreditCardExpiration = new System.Windows.Forms.TextBox();
			this.LblCreditCardExpiration = new System.Windows.Forms.Label();
			this.TxtBillingCycleDate = new System.Windows.Forms.TextBox();
			this.LblBillingCycleDate = new System.Windows.Forms.Label();
			this.TxtAmountPastDue = new System.Windows.Forms.TextBox();
			this.LblAmountPastDue = new System.Windows.Forms.Label();
			this.LblPayMethod = new System.Windows.Forms.Label();
			this.LblAccBalance = new System.Windows.Forms.Label();
			this.TxtPayMethod = new System.Windows.Forms.TextBox();
			this.TxtAccBalance = new System.Windows.Forms.TextBox();
			this.TxtAccState = new System.Windows.Forms.TextBox();
			this.LblAccState = new System.Windows.Forms.Label();
			this.BtnGetAccount = new System.Windows.Forms.Button();
			this.OptLastInvoice = new System.Windows.Forms.RadioButton();
			this.OptInvoiceDate = new System.Windows.Forms.RadioButton();
			this.OptRange = new System.Windows.Forms.RadioButton();
			this.label1 = new System.Windows.Forms.Label();
			this.DatePickInvoice = new System.Windows.Forms.DateTimePicker();
			this.DatePickFrom = new System.Windows.Forms.DateTimePicker();
			this.DatePickTo = new System.Windows.Forms.DateTimePicker();
			this.TxtEntriesPerPage = new System.Windows.Forms.TextBox();
			this.LblEntryCount = new System.Windows.Forms.Label();
			this.LblPageNumber = new System.Windows.Forms.Label();
			this.TxtPageNumber = new System.Windows.Forms.TextBox();
			this.LblTotalPages = new System.Windows.Forms.Label();
			this.TxtTotalNumberOfPages = new System.Windows.Forms.TextBox();
			this.TxtLastPaymentDate = new System.Windows.Forms.TextBox();
			this.LblLastPaymentDate = new System.Windows.Forms.Label();
			this.BtnNextPage = new System.Windows.Forms.Button();
			this.GrpResults.SuspendLayout();
			this.GrpAccountEntry.SuspendLayout();
			this.SuspendLayout();
			// 
			// GrpResults
			// 
			this.GrpResults.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.TxtLastPaymentDate,
																					 this.LblLastPaymentDate,
																					 this.GrpAccountEntry,
																					 this.TxtAccountID,
																					 this.LblAccountID,
																					 this.TxtPastDue,
																					 this.LblPastDue,
																					 this.LblLastAmountPaid,
																					 this.TxtLastAmountPaid,
																					 this.TxtCreditCardModifyDate,
																					 this.LblCreditCardModifyDate,
																					 this.TxtCreditCardInfo,
																					 this.LblCreditCardInfo,
																					 this.TxtCreditCardExpiration,
																					 this.LblCreditCardExpiration,
																					 this.TxtBillingCycleDate,
																					 this.LblBillingCycleDate,
																					 this.TxtAmountPastDue,
																					 this.LblAmountPastDue,
																					 this.LblPayMethod,
																					 this.LblAccBalance,
																					 this.TxtPayMethod,
																					 this.TxtAccBalance,
																					 this.TxtAccState,
																					 this.LblAccState});
			this.GrpResults.Location = new System.Drawing.Point(8, 128);
			this.GrpResults.Name = "GrpResults";
			this.GrpResults.Size = new System.Drawing.Size(760, 456);
			this.GrpResults.TabIndex = 45;
			this.GrpResults.TabStop = false;
			this.GrpResults.Text = "Results";
			// 
			// GrpAccountEntry
			// 
			this.GrpAccountEntry.Controls.AddRange(new System.Windows.Forms.Control[] {
																						  this.BtnNextPage,
																						  this.LstAccountEntries,
																						  this.LblTotalPages,
																						  this.TxtTotalNumberOfPages,
																						  this.LblPageNumber,
																						  this.TxtPageNumber,
																						  this.LblEntryCount,
																						  this.TxtEntriesPerPage});
			this.GrpAccountEntry.Location = new System.Drawing.Point(16, 152);
			this.GrpAccountEntry.Name = "GrpAccountEntry";
			this.GrpAccountEntry.Size = new System.Drawing.Size(728, 288);
			this.GrpAccountEntry.TabIndex = 38;
			this.GrpAccountEntry.TabStop = false;
			this.GrpAccountEntry.Text = "Account Entries";
			// 
			// LstAccountEntries
			// 
			this.LstAccountEntries.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																								this.ClmRef,
																								this.ClmType,
																								this.ClmItemId,
																								this.ClmDescription,
																								this.ClmNetAmount,
																								this.ClmGrossAmount,
																								this.ClmVAT,
																								this.ClmBalance,
																								this.ClmDate});
			this.LstAccountEntries.GridLines = true;
			this.LstAccountEntries.Location = new System.Drawing.Point(16, 56);
			this.LstAccountEntries.Name = "LstAccountEntries";
			this.LstAccountEntries.Size = new System.Drawing.Size(688, 224);
			this.LstAccountEntries.TabIndex = 15;
			this.LstAccountEntries.View = System.Windows.Forms.View.Details;
			// 
			// ClmRef
			// 
			this.ClmRef.Text = "Reference";
			this.ClmRef.Width = 70;
			// 
			// ClmType
			// 
			this.ClmType.Text = "Type";
			this.ClmType.Width = 100;
			// 
			// ClmItemId
			// 
			this.ClmItemId.Text = "Item Id";
			this.ClmItemId.Width = 74;
			// 
			// ClmDescription
			// 
			this.ClmDescription.Text = "Description";
			this.ClmDescription.Width = 80;
			// 
			// ClmNetAmount
			// 
			this.ClmNetAmount.Text = "NetAmount";
			this.ClmNetAmount.Width = 80;
			// 
			// ClmGrossAmount
			// 
			this.ClmGrossAmount.Text = "GrossAmount";
			this.ClmGrossAmount.Width = 80;
			// 
			// ClmVAT
			// 
			this.ClmVAT.Text = "VAT";
			this.ClmVAT.Width = 40;
			// 
			// ClmBalance
			// 
			this.ClmBalance.Text = "Balance";
			// 
			// ClmDate
			// 
			this.ClmDate.Text = "Date";
			this.ClmDate.Width = 98;
			// 
			// TxtAccountID
			// 
			this.TxtAccountID.Location = new System.Drawing.Point(128, 24);
			this.TxtAccountID.Name = "TxtAccountID";
			this.TxtAccountID.ReadOnly = true;
			this.TxtAccountID.Size = new System.Drawing.Size(104, 20);
			this.TxtAccountID.TabIndex = 37;
			this.TxtAccountID.Text = "";
			// 
			// LblAccountID
			// 
			this.LblAccountID.Location = new System.Drawing.Point(16, 24);
			this.LblAccountID.Name = "LblAccountID";
			this.LblAccountID.Size = new System.Drawing.Size(104, 23);
			this.LblAccountID.TabIndex = 36;
			this.LblAccountID.Text = "Account ID:";
			// 
			// TxtPastDue
			// 
			this.TxtPastDue.Location = new System.Drawing.Point(360, 88);
			this.TxtPastDue.Name = "TxtPastDue";
			this.TxtPastDue.ReadOnly = true;
			this.TxtPastDue.Size = new System.Drawing.Size(104, 20);
			this.TxtPastDue.TabIndex = 35;
			this.TxtPastDue.Text = "";
			// 
			// LblPastDue
			// 
			this.LblPastDue.Location = new System.Drawing.Point(248, 88);
			this.LblPastDue.Name = "LblPastDue";
			this.LblPastDue.Size = new System.Drawing.Size(64, 23);
			this.LblPastDue.TabIndex = 34;
			this.LblPastDue.Text = "Past Due:";
			// 
			// LblLastAmountPaid
			// 
			this.LblLastAmountPaid.Location = new System.Drawing.Point(16, 56);
			this.LblLastAmountPaid.Name = "LblLastAmountPaid";
			this.LblLastAmountPaid.Size = new System.Drawing.Size(104, 23);
			this.LblLastAmountPaid.TabIndex = 33;
			this.LblLastAmountPaid.Text = "Last Amount Paid:";
			// 
			// TxtLastAmountPaid
			// 
			this.TxtLastAmountPaid.Location = new System.Drawing.Point(128, 56);
			this.TxtLastAmountPaid.Name = "TxtLastAmountPaid";
			this.TxtLastAmountPaid.ReadOnly = true;
			this.TxtLastAmountPaid.Size = new System.Drawing.Size(104, 20);
			this.TxtLastAmountPaid.TabIndex = 32;
			this.TxtLastAmountPaid.Text = "";
			// 
			// TxtCreditCardModifyDate
			// 
			this.TxtCreditCardModifyDate.Location = new System.Drawing.Point(608, 120);
			this.TxtCreditCardModifyDate.Name = "TxtCreditCardModifyDate";
			this.TxtCreditCardModifyDate.ReadOnly = true;
			this.TxtCreditCardModifyDate.Size = new System.Drawing.Size(104, 20);
			this.TxtCreditCardModifyDate.TabIndex = 31;
			this.TxtCreditCardModifyDate.Text = "";
			// 
			// LblCreditCardModifyDate
			// 
			this.LblCreditCardModifyDate.Location = new System.Drawing.Point(480, 120);
			this.LblCreditCardModifyDate.Name = "LblCreditCardModifyDate";
			this.LblCreditCardModifyDate.Size = new System.Drawing.Size(136, 23);
			this.LblCreditCardModifyDate.TabIndex = 30;
			this.LblCreditCardModifyDate.Text = "Credit Card Modify Date:";
			// 
			// TxtCreditCardInfo
			// 
			this.TxtCreditCardInfo.Location = new System.Drawing.Point(128, 120);
			this.TxtCreditCardInfo.Name = "TxtCreditCardInfo";
			this.TxtCreditCardInfo.ReadOnly = true;
			this.TxtCreditCardInfo.Size = new System.Drawing.Size(104, 20);
			this.TxtCreditCardInfo.TabIndex = 29;
			this.TxtCreditCardInfo.Text = "";
			// 
			// LblCreditCardInfo
			// 
			this.LblCreditCardInfo.Location = new System.Drawing.Point(16, 120);
			this.LblCreditCardInfo.Name = "LblCreditCardInfo";
			this.LblCreditCardInfo.Size = new System.Drawing.Size(104, 23);
			this.LblCreditCardInfo.TabIndex = 28;
			this.LblCreditCardInfo.Text = "Credit Card Info:";
			// 
			// TxtCreditCardExpiration
			// 
			this.TxtCreditCardExpiration.Location = new System.Drawing.Point(360, 120);
			this.TxtCreditCardExpiration.Name = "TxtCreditCardExpiration";
			this.TxtCreditCardExpiration.ReadOnly = true;
			this.TxtCreditCardExpiration.Size = new System.Drawing.Size(104, 20);
			this.TxtCreditCardExpiration.TabIndex = 27;
			this.TxtCreditCardExpiration.Text = "";
			// 
			// LblCreditCardExpiration
			// 
			this.LblCreditCardExpiration.Location = new System.Drawing.Point(248, 120);
			this.LblCreditCardExpiration.Name = "LblCreditCardExpiration";
			this.LblCreditCardExpiration.Size = new System.Drawing.Size(136, 23);
			this.LblCreditCardExpiration.TabIndex = 26;
			this.LblCreditCardExpiration.Text = "Credit Card Expiration:";
			// 
			// TxtBillingCycleDate
			// 
			this.TxtBillingCycleDate.Location = new System.Drawing.Point(608, 24);
			this.TxtBillingCycleDate.Name = "TxtBillingCycleDate";
			this.TxtBillingCycleDate.ReadOnly = true;
			this.TxtBillingCycleDate.Size = new System.Drawing.Size(104, 20);
			this.TxtBillingCycleDate.TabIndex = 25;
			this.TxtBillingCycleDate.Text = "";
			// 
			// LblBillingCycleDate
			// 
			this.LblBillingCycleDate.Location = new System.Drawing.Point(480, 24);
			this.LblBillingCycleDate.Name = "LblBillingCycleDate";
			this.LblBillingCycleDate.Size = new System.Drawing.Size(96, 23);
			this.LblBillingCycleDate.TabIndex = 24;
			this.LblBillingCycleDate.Text = "Billing Cycle Date:";
			// 
			// TxtAmountPastDue
			// 
			this.TxtAmountPastDue.Location = new System.Drawing.Point(128, 88);
			this.TxtAmountPastDue.Name = "TxtAmountPastDue";
			this.TxtAmountPastDue.ReadOnly = true;
			this.TxtAmountPastDue.Size = new System.Drawing.Size(104, 20);
			this.TxtAmountPastDue.TabIndex = 23;
			this.TxtAmountPastDue.Text = "";
			// 
			// LblAmountPastDue
			// 
			this.LblAmountPastDue.Location = new System.Drawing.Point(16, 88);
			this.LblAmountPastDue.Name = "LblAmountPastDue";
			this.LblAmountPastDue.Size = new System.Drawing.Size(104, 23);
			this.LblAmountPastDue.TabIndex = 22;
			this.LblAmountPastDue.Text = "Amount Past Due:";
			// 
			// LblPayMethod
			// 
			this.LblPayMethod.Location = new System.Drawing.Point(480, 56);
			this.LblPayMethod.Name = "LblPayMethod";
			this.LblPayMethod.Size = new System.Drawing.Size(96, 23);
			this.LblPayMethod.TabIndex = 21;
			this.LblPayMethod.Text = "Payment Method::";
			// 
			// LblAccBalance
			// 
			this.LblAccBalance.Location = new System.Drawing.Point(248, 56);
			this.LblAccBalance.Name = "LblAccBalance";
			this.LblAccBalance.Size = new System.Drawing.Size(96, 23);
			this.LblAccBalance.TabIndex = 20;
			this.LblAccBalance.Text = "Current Balance:";
			// 
			// TxtPayMethod
			// 
			this.TxtPayMethod.Location = new System.Drawing.Point(608, 56);
			this.TxtPayMethod.Name = "TxtPayMethod";
			this.TxtPayMethod.ReadOnly = true;
			this.TxtPayMethod.Size = new System.Drawing.Size(104, 20);
			this.TxtPayMethod.TabIndex = 19;
			this.TxtPayMethod.Text = "";
			// 
			// TxtAccBalance
			// 
			this.TxtAccBalance.Location = new System.Drawing.Point(360, 56);
			this.TxtAccBalance.Name = "TxtAccBalance";
			this.TxtAccBalance.ReadOnly = true;
			this.TxtAccBalance.Size = new System.Drawing.Size(104, 20);
			this.TxtAccBalance.TabIndex = 18;
			this.TxtAccBalance.Text = "";
			// 
			// TxtAccState
			// 
			this.TxtAccState.Location = new System.Drawing.Point(360, 24);
			this.TxtAccState.Name = "TxtAccState";
			this.TxtAccState.ReadOnly = true;
			this.TxtAccState.Size = new System.Drawing.Size(104, 20);
			this.TxtAccState.TabIndex = 17;
			this.TxtAccState.Text = "";
			// 
			// LblAccState
			// 
			this.LblAccState.Location = new System.Drawing.Point(248, 24);
			this.LblAccState.Name = "LblAccState";
			this.LblAccState.Size = new System.Drawing.Size(88, 23);
			this.LblAccState.TabIndex = 16;
			this.LblAccState.Text = "Account State:";
			// 
			// BtnGetAccount
			// 
			this.BtnGetAccount.Location = new System.Drawing.Point(160, 88);
			this.BtnGetAccount.Name = "BtnGetAccount";
			this.BtnGetAccount.Size = new System.Drawing.Size(120, 26);
			this.BtnGetAccount.TabIndex = 46;
			this.BtnGetAccount.Text = "GetAccount";
			this.BtnGetAccount.Click += new System.EventHandler(this.BtnGetAccount_Click);
			// 
			// OptLastInvoice
			// 
			this.OptLastInvoice.Location = new System.Drawing.Point(16, 8);
			this.OptLastInvoice.Name = "OptLastInvoice";
			this.OptLastInvoice.Size = new System.Drawing.Size(128, 24);
			this.OptLastInvoice.TabIndex = 47;
			this.OptLastInvoice.Text = "Since Last Invoice";
			this.OptLastInvoice.CheckedChanged += new System.EventHandler(this.Option_CheckedChanged);
			// 
			// OptInvoiceDate
			// 
			this.OptInvoiceDate.Location = new System.Drawing.Point(16, 32);
			this.OptInvoiceDate.Name = "OptInvoiceDate";
			this.OptInvoiceDate.Size = new System.Drawing.Size(128, 24);
			this.OptInvoiceDate.TabIndex = 48;
			this.OptInvoiceDate.Text = "Invoice Date";
			this.OptInvoiceDate.CheckedChanged += new System.EventHandler(this.Option_CheckedChanged);
			// 
			// OptRange
			// 
			this.OptRange.Location = new System.Drawing.Point(16, 56);
			this.OptRange.Name = "OptRange";
			this.OptRange.Size = new System.Drawing.Size(128, 24);
			this.OptRange.TabIndex = 49;
			this.OptRange.Text = "Between Dates";
			this.OptRange.CheckedChanged += new System.EventHandler(this.Option_CheckedChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(304, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(16, 23);
			this.label1.TabIndex = 54;
			this.label1.Text = " - ";
			// 
			// DatePickInvoice
			// 
			this.DatePickInvoice.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickInvoice.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickInvoice.Location = new System.Drawing.Point(152, 32);
			this.DatePickInvoice.Name = "DatePickInvoice";
			this.DatePickInvoice.Size = new System.Drawing.Size(152, 20);
			this.DatePickInvoice.TabIndex = 55;
			// 
			// DatePickFrom
			// 
			this.DatePickFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickFrom.Location = new System.Drawing.Point(152, 56);
			this.DatePickFrom.Name = "DatePickFrom";
			this.DatePickFrom.Size = new System.Drawing.Size(152, 20);
			this.DatePickFrom.TabIndex = 56;
			// 
			// DatePickTo
			// 
			this.DatePickTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickTo.Location = new System.Drawing.Point(320, 56);
			this.DatePickTo.Name = "DatePickTo";
			this.DatePickTo.Size = new System.Drawing.Size(152, 20);
			this.DatePickTo.TabIndex = 57;
			// 
			// TxtEntriesPerPage
			// 
			this.TxtEntriesPerPage.Location = new System.Drawing.Point(320, 25);
			this.TxtEntriesPerPage.Name = "TxtEntriesPerPage";
			this.TxtEntriesPerPage.ReadOnly = true;
			this.TxtEntriesPerPage.Size = new System.Drawing.Size(104, 20);
			this.TxtEntriesPerPage.TabIndex = 59;
			this.TxtEntriesPerPage.Text = "";
			// 
			// LblEntryCount
			// 
			this.LblEntryCount.Location = new System.Drawing.Point(248, 24);
			this.LblEntryCount.Name = "LblEntryCount";
			this.LblEntryCount.Size = new System.Drawing.Size(72, 23);
			this.LblEntryCount.TabIndex = 58;
			this.LblEntryCount.Text = "Entry Count:";
			// 
			// LblPageNumber
			// 
			this.LblPageNumber.Location = new System.Drawing.Point(16, 24);
			this.LblPageNumber.Name = "LblPageNumber";
			this.LblPageNumber.Size = new System.Drawing.Size(32, 23);
			this.LblPageNumber.TabIndex = 60;
			this.LblPageNumber.Text = "Page";
			// 
			// TxtPageNumber
			// 
			this.TxtPageNumber.Location = new System.Drawing.Point(56, 25);
			this.TxtPageNumber.Name = "TxtPageNumber";
			this.TxtPageNumber.ReadOnly = true;
			this.TxtPageNumber.Size = new System.Drawing.Size(32, 20);
			this.TxtPageNumber.TabIndex = 61;
			this.TxtPageNumber.Text = "";
			// 
			// LblTotalPages
			// 
			this.LblTotalPages.Location = new System.Drawing.Point(96, 24);
			this.LblTotalPages.Name = "LblTotalPages";
			this.LblTotalPages.Size = new System.Drawing.Size(24, 23);
			this.LblTotalPages.TabIndex = 62;
			this.LblTotalPages.Text = "of ";
			// 
			// TxtTotalNumberOfPages
			// 
			this.TxtTotalNumberOfPages.Location = new System.Drawing.Point(120, 25);
			this.TxtTotalNumberOfPages.Name = "TxtTotalNumberOfPages";
			this.TxtTotalNumberOfPages.ReadOnly = true;
			this.TxtTotalNumberOfPages.Size = new System.Drawing.Size(48, 20);
			this.TxtTotalNumberOfPages.TabIndex = 63;
			this.TxtTotalNumberOfPages.Text = "";
			// 
			// TxtLastPaymentDate
			// 
			this.TxtLastPaymentDate.Location = new System.Drawing.Point(608, 88);
			this.TxtLastPaymentDate.Name = "TxtLastPaymentDate";
			this.TxtLastPaymentDate.ReadOnly = true;
			this.TxtLastPaymentDate.Size = new System.Drawing.Size(104, 20);
			this.TxtLastPaymentDate.TabIndex = 40;
			this.TxtLastPaymentDate.Text = "";
			// 
			// LblLastPaymentDate
			// 
			this.LblLastPaymentDate.Location = new System.Drawing.Point(480, 88);
			this.LblLastPaymentDate.Name = "LblLastPaymentDate";
			this.LblLastPaymentDate.Size = new System.Drawing.Size(112, 23);
			this.LblLastPaymentDate.TabIndex = 39;
			this.LblLastPaymentDate.Text = "Last Payment Date:";
			// 
			// BtnNextPage
			// 
			this.BtnNextPage.Location = new System.Drawing.Point(448, 24);
			this.BtnNextPage.Name = "BtnNextPage";
			this.BtnNextPage.TabIndex = 64;
			this.BtnNextPage.Text = "Next Page";
			this.BtnNextPage.Visible = false;
			this.BtnNextPage.Click += new System.EventHandler(this.BtnNextPage_Click);
			// 
			// FrmGetAccount
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(792, 593);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.DatePickTo,
																		  this.DatePickFrom,
																		  this.DatePickInvoice,
																		  this.label1,
																		  this.OptRange,
																		  this.OptInvoiceDate,
																		  this.OptLastInvoice,
																		  this.BtnGetAccount,
																		  this.GrpResults});
			this.Name = "FrmGetAccount";
			this.Text = "GetAccount";
			this.Load += new System.EventHandler(this.FrmGetAccount_Load);
			this.GrpResults.ResumeLayout(false);
			this.GrpAccountEntry.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
	
		private void Option_CheckedChanged(object sender, System.EventArgs e)
		{
			RadioButton sel = (RadioButton) sender;
			if (sel.Checked &&  sel == this.OptLastInvoice)
			{
				DatePickInvoice.Enabled = false;
				DatePickFrom.Enabled = false;
				DatePickTo.Enabled = false;
			} 
			else if (sel.Checked && sel == OptInvoiceDate)
			{
				DatePickInvoice.Enabled = true;
				DatePickFrom.Enabled = false;
				DatePickTo.Enabled = false;
				

			}
			else if (sel.Checked && sel == OptRange)
			{ 
				DatePickInvoice.Enabled = false;
				DatePickFrom.Enabled = true;
				DatePickTo.Enabled = true;
			}
		}

		private void FrmGetAccount_Load(object sender, System.EventArgs e)
		{
			OptRange.Checked = true;
			DateTime now = DateTime.Now;
			DatePickInvoice.Value = now;
			DatePickTo.Value = now;
			DatePickFrom.Value = now.AddDays(-5);

		}

		private void BtnGetAccount_Click(object sender, System.EventArgs e)
		{
			CallGetAccount(1);
		}

		private void CallGetAccount(int PageNumber)
		{
			try
			{
				LstAccountEntries.Items.Clear();
				TxtAccState.Text = "";
				TxtBillingCycleDate.Text = "";
				TxtLastAmountPaid.Text = "";
				TxtAccBalance.Text = "";
				TxtPayMethod.Text = "";
				TxtAmountPastDue.Text = "";
				TxtPastDue.Text = "";
				TxtCreditCardInfo.Text = "";
				TxtCreditCardExpiration.Text = "";
				TxtCreditCardModifyDate.Text = "";
				TxtAccountID.Text = "";
				TxtLastPaymentDate.Text = "";
				TxtEntriesPerPage.Text = "";
				TxtPageNumber.Text = "";
				TxtTotalNumberOfPages.Text = "";
	
				GetAccountCall apicall = new GetAccountCall(Context);
			
				AccountHistorySelectionCodeType	ahselect = AccountHistorySelectionCodeType.CustomCode;
				if (OptLastInvoice.Checked == true)
				{
					ahselect = AccountHistorySelectionCodeType.LastInvoice;

				} 
				else if (OptInvoiceDate.Checked == true)
				{
					ahselect = AccountHistorySelectionCodeType.SpecifiedInvoice;
					apicall.InvoiceDate = DatePickInvoice.Value;
				}
				else if (OptRange.Checked == true)
				{
					ahselect = AccountHistorySelectionCodeType.BetweenSpecifiedDates;
					apicall.StartTimeFilter = new TimeFilter(DatePickFrom.Value, DatePickTo.Value);

				}
			
				apicall.Pagination = new PaginationType();
				apicall.Pagination.PageNumber = PageNumber;

				//apicall.ApiRequest.ExcludeSummary = true;
				//apicall.ApiRequest.Pagination.EntriesPerPage = 1;
				AccountEntryTypeCollection entrylist = apicall.GetAccount(ahselect);
	
				TxtAccountID.Text = apicall.AccountID;
				if (apicall.AccountSummary != null)
				{
					AccountSummaryType Summary = apicall.AccountSummary;
					TxtAccState.Text = Summary.AccountState.ToString();
					TxtBillingCycleDate.Text = Summary.BillingCycleDate.ToString();
					TxtLastAmountPaid.Text = Summary.LastAmountPaid.Value.ToString();
					TxtAccBalance.Text = Summary.CurrentBalance.Value.ToString();
					TxtPayMethod.Text = Summary.PaymentMethod.ToString();
					TxtAmountPastDue.Text = Summary.AmountPastDue.Value.ToString();
					TxtPastDue.Text = Summary.PastDue.ToString();
					TxtCreditCardInfo.Text = Summary.CreditCardInfo;
					if (Summary.CreditCardExpirationSpecified)
						TxtCreditCardExpiration.Text = Summary.CreditCardExpiration.ToString();

					if (Summary.CreditCardModifyDateSpecified)
						TxtCreditCardModifyDate.Text = Summary.CreditCardModifyDate.ToString();

					if (Summary.LastPaymentDateSpecified)
						TxtLastPaymentDate.Text = Summary.LastPaymentDate.ToString();
				/*	Summary.InvoiceBalance;
					Summary.InvoiceCredit;
					Summary.InvoiceDate;
					Summary.InvoiceNewFee;
					Summary.InvoicePayment;
					*/
					TxtEntriesPerPage.Text = apicall.ApiResponse.EntriesPerPage.ToString();
					TxtPageNumber.Text = apicall.ApiResponse.PageNumber.ToString();
					TxtTotalNumberOfPages.Text = apicall.ApiResponse.PaginationResult.TotalNumberOfPages.ToString();

				}

				if (TxtPageNumber.Text.Length > 0 && Int32.Parse(TxtPageNumber.Text) < Int32.Parse(TxtTotalNumberOfPages.Text))
					BtnNextPage.Visible = true;;

		
				if (entrylist == null)
					return;

				foreach (AccountEntryType entry in entrylist)
				{
					string[] listparams = new string[9];
					listparams[0] = entry.RefNumber;
					listparams[1] = entry.AccountDetailsEntryType.ToString();
					listparams[2] = entry.ItemID;
					listparams[3] = entry.Description;
					if (entry.NetDetailAmount != null)
						listparams[4] = entry.NetDetailAmount.Value.ToString();

					if (entry.GrossDetailAmount != null)
						listparams[5] = entry.GrossDetailAmount.Value.ToString();
					listparams[6] = entry.VATPercent.ToString();
					listparams[7] = entry.Balance.Value.ToString();
					listparams[8] = entry.Date.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					this.LstAccountEntries.Items.Add(vi);
				}
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void BtnNextPage_Click(object sender, System.EventArgs e)
		{
			
			int pagenumber = Int32.Parse(TxtPageNumber.Text) + 1;
			CallGetAccount(pagenumber);
		}
	}
}
