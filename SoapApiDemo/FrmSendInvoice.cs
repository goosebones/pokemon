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
	/// Summary description for GetUserForm.
	/// </summary>
	public class FrmSendInvoice : System.Windows.Forms.Form
	{
		public ApiContext Context;
		public string mItemID;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.ComboBox CboInsuranceOption;
		private System.Windows.Forms.Label LblInsuranceOption;
		private System.Windows.Forms.TextBox TxtTransactionId;
		private System.Windows.Forms.Label LblTransactionId;
		private System.Windows.Forms.TextBox TxtInsuranceFee;
		private System.Windows.Forms.Label lblInsuranceFee;
		private System.Windows.Forms.Label LblPaymentMethod;
		private System.Windows.Forms.ComboBox CboPaymentMethod;
		private System.Windows.Forms.TextBox TxtPayPalEmailAddress;
		private System.Windows.Forms.Label LblPayPalEmailAddress;
		private System.Windows.Forms.TextBox TxtCheckoutInstructions;
		private System.Windows.Forms.Label LblCheckoutInstructions;
		private System.Windows.Forms.CheckBox ChkEmailCopyToSeller;
		private System.Windows.Forms.Button BtnSendInvoice;
		private System.Windows.Forms.TextBox TxtShippingInsuranceCost;
		private System.Windows.Forms.Label LblShippingInsuranceCost;
		private System.Windows.Forms.ComboBox CboShippingService;
		private System.Windows.Forms.Label LblShippingService;
		private System.Windows.Forms.Label LblShippingServiceCost;
		private System.Windows.Forms.Label LblShippingServiceAdditionalCost;
		private System.Windows.Forms.Label LblShippingServicePriority;
		private System.Windows.Forms.TextBox TxtShippingServicePriority;
		private System.Windows.Forms.GroupBox GrpShippingServiceOptions;
		private System.Windows.Forms.GroupBox GrpSaleTaxType;
		private System.Windows.Forms.TextBox TxtSalesTaxPercent;
		private System.Windows.Forms.Label LblSalesTaxPercent;
		private System.Windows.Forms.TextBox TxtSalesTaxState;
		private System.Windows.Forms.Label LblSalesTaxState;
		private System.Windows.Forms.CheckBox ChkShippingIncludedInTax;
		private System.Windows.Forms.TextBox TxtSalesTaxAmount;
		private System.Windows.Forms.Label LblSalesTaxAmount;
		private System.Windows.Forms.TextBox TxtShippingServiceCost;
		private System.Windows.Forms.TextBox TxtShippingServiceAdditionalCost;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmSendInvoice()
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
			this.BtnSendInvoice = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.CboInsuranceOption = new System.Windows.Forms.ComboBox();
			this.LblInsuranceOption = new System.Windows.Forms.Label();
			this.TxtTransactionId = new System.Windows.Forms.TextBox();
			this.LblTransactionId = new System.Windows.Forms.Label();
			this.TxtInsuranceFee = new System.Windows.Forms.TextBox();
			this.lblInsuranceFee = new System.Windows.Forms.Label();
			this.CboPaymentMethod = new System.Windows.Forms.ComboBox();
			this.LblPaymentMethod = new System.Windows.Forms.Label();
			this.TxtPayPalEmailAddress = new System.Windows.Forms.TextBox();
			this.LblPayPalEmailAddress = new System.Windows.Forms.Label();
			this.TxtCheckoutInstructions = new System.Windows.Forms.TextBox();
			this.LblCheckoutInstructions = new System.Windows.Forms.Label();
			this.ChkEmailCopyToSeller = new System.Windows.Forms.CheckBox();
			this.TxtShippingInsuranceCost = new System.Windows.Forms.TextBox();
			this.LblShippingInsuranceCost = new System.Windows.Forms.Label();
			this.CboShippingService = new System.Windows.Forms.ComboBox();
			this.LblShippingService = new System.Windows.Forms.Label();
			this.TxtShippingServiceCost = new System.Windows.Forms.TextBox();
			this.LblShippingServiceCost = new System.Windows.Forms.Label();
			this.TxtShippingServiceAdditionalCost = new System.Windows.Forms.TextBox();
			this.LblShippingServiceAdditionalCost = new System.Windows.Forms.Label();
			this.TxtShippingServicePriority = new System.Windows.Forms.TextBox();
			this.LblShippingServicePriority = new System.Windows.Forms.Label();
			this.GrpShippingServiceOptions = new System.Windows.Forms.GroupBox();
			this.GrpSaleTaxType = new System.Windows.Forms.GroupBox();
			this.TxtSalesTaxAmount = new System.Windows.Forms.TextBox();
			this.LblSalesTaxAmount = new System.Windows.Forms.Label();
			this.ChkShippingIncludedInTax = new System.Windows.Forms.CheckBox();
			this.TxtSalesTaxState = new System.Windows.Forms.TextBox();
			this.LblSalesTaxState = new System.Windows.Forms.Label();
			this.TxtSalesTaxPercent = new System.Windows.Forms.TextBox();
			this.LblSalesTaxPercent = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.GrpShippingServiceOptions.SuspendLayout();
			this.GrpSaleTaxType.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnSendInvoice
			// 
			this.BtnSendInvoice.Location = new System.Drawing.Point(136, 560);
			this.BtnSendInvoice.Name = "BtnSendInvoice";
			this.BtnSendInvoice.Size = new System.Drawing.Size(120, 23);
			this.BtnSendInvoice.TabIndex = 23;
			this.BtnSendInvoice.Text = "SendInvoice";
			this.BtnSendInvoice.Click += new System.EventHandler(this.BtnSendInvoice_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblStatus,
																					this.TxtStatus});
			this.GrpResult.Location = new System.Drawing.Point(16, 592);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(352, 64);
			this.GrpResult.TabIndex = 25;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblStatus
			// 
			this.LblStatus.Location = new System.Drawing.Point(16, 24);
			this.LblStatus.Name = "LblStatus";
			this.LblStatus.Size = new System.Drawing.Size(80, 23);
			this.LblStatus.TabIndex = 42;
			this.LblStatus.Text = "Status:";
			// 
			// TxtStatus
			// 
			this.TxtStatus.Location = new System.Drawing.Point(96, 24);
			this.TxtStatus.Name = "TxtStatus";
			this.TxtStatus.ReadOnly = true;
			this.TxtStatus.Size = new System.Drawing.Size(72, 20);
			this.TxtStatus.TabIndex = 41;
			this.TxtStatus.Text = "";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(80, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(104, 20);
			this.TxtItemId.TabIndex = 26;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(24, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(48, 23);
			this.LblItemId.TabIndex = 27;
			this.LblItemId.Text = "Item Id:";
			// 
			// CboInsuranceOption
			// 
			this.CboInsuranceOption.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboInsuranceOption.Location = new System.Drawing.Point(120, 392);
			this.CboInsuranceOption.Name = "CboInsuranceOption";
			this.CboInsuranceOption.Size = new System.Drawing.Size(144, 21);
			this.CboInsuranceOption.TabIndex = 55;
			// 
			// LblInsuranceOption
			// 
			this.LblInsuranceOption.Location = new System.Drawing.Point(24, 392);
			this.LblInsuranceOption.Name = "LblInsuranceOption";
			this.LblInsuranceOption.Size = new System.Drawing.Size(96, 18);
			this.LblInsuranceOption.TabIndex = 54;
			this.LblInsuranceOption.Text = "Insurance Option:";
			// 
			// TxtTransactionId
			// 
			this.TxtTransactionId.Location = new System.Drawing.Point(288, 8);
			this.TxtTransactionId.Name = "TxtTransactionId";
			this.TxtTransactionId.Size = new System.Drawing.Size(112, 20);
			this.TxtTransactionId.TabIndex = 56;
			this.TxtTransactionId.Text = "";
			// 
			// LblTransactionId
			// 
			this.LblTransactionId.Location = new System.Drawing.Point(192, 8);
			this.LblTransactionId.Name = "LblTransactionId";
			this.LblTransactionId.Size = new System.Drawing.Size(80, 23);
			this.LblTransactionId.TabIndex = 57;
			this.LblTransactionId.Text = "Transaction Id:";
			// 
			// TxtInsuranceFee
			// 
			this.TxtInsuranceFee.Location = new System.Drawing.Point(352, 392);
			this.TxtInsuranceFee.Name = "TxtInsuranceFee";
			this.TxtInsuranceFee.Size = new System.Drawing.Size(48, 20);
			this.TxtInsuranceFee.TabIndex = 58;
			this.TxtInsuranceFee.Text = "3.14";
			// 
			// lblInsuranceFee
			// 
			this.lblInsuranceFee.Location = new System.Drawing.Point(272, 392);
			this.lblInsuranceFee.Name = "lblInsuranceFee";
			this.lblInsuranceFee.Size = new System.Drawing.Size(80, 23);
			this.lblInsuranceFee.TabIndex = 59;
			this.lblInsuranceFee.Text = "Insurance Fee:";
			// 
			// CboPaymentMethod
			// 
			this.CboPaymentMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboPaymentMethod.Location = new System.Drawing.Point(160, 424);
			this.CboPaymentMethod.Name = "CboPaymentMethod";
			this.CboPaymentMethod.Size = new System.Drawing.Size(240, 21);
			this.CboPaymentMethod.TabIndex = 61;
			// 
			// LblPaymentMethod
			// 
			this.LblPaymentMethod.Location = new System.Drawing.Point(24, 424);
			this.LblPaymentMethod.Name = "LblPaymentMethod";
			this.LblPaymentMethod.Size = new System.Drawing.Size(96, 18);
			this.LblPaymentMethod.TabIndex = 60;
			this.LblPaymentMethod.Text = "Payment Method:";
			// 
			// TxtPayPalEmailAddress
			// 
			this.TxtPayPalEmailAddress.Location = new System.Drawing.Point(160, 456);
			this.TxtPayPalEmailAddress.Name = "TxtPayPalEmailAddress";
			this.TxtPayPalEmailAddress.Size = new System.Drawing.Size(240, 20);
			this.TxtPayPalEmailAddress.TabIndex = 65;
			this.TxtPayPalEmailAddress.Text = "test@my.com";
			// 
			// LblPayPalEmailAddress
			// 
			this.LblPayPalEmailAddress.Location = new System.Drawing.Point(24, 456);
			this.LblPayPalEmailAddress.Name = "LblPayPalEmailAddress";
			this.LblPayPalEmailAddress.Size = new System.Drawing.Size(120, 18);
			this.LblPayPalEmailAddress.TabIndex = 66;
			this.LblPayPalEmailAddress.Text = "PayPal Email Address:";
			// 
			// TxtCheckoutInstructions
			// 
			this.TxtCheckoutInstructions.Location = new System.Drawing.Point(160, 488);
			this.TxtCheckoutInstructions.Name = "TxtCheckoutInstructions";
			this.TxtCheckoutInstructions.Size = new System.Drawing.Size(240, 20);
			this.TxtCheckoutInstructions.TabIndex = 67;
			this.TxtCheckoutInstructions.Text = "";
			// 
			// LblCheckoutInstructions
			// 
			this.LblCheckoutInstructions.Location = new System.Drawing.Point(24, 488);
			this.LblCheckoutInstructions.Name = "LblCheckoutInstructions";
			this.LblCheckoutInstructions.Size = new System.Drawing.Size(120, 18);
			this.LblCheckoutInstructions.TabIndex = 68;
			this.LblCheckoutInstructions.Text = "Checkout Instructions:";
			// 
			// ChkEmailCopyToSeller
			// 
			this.ChkEmailCopyToSeller.Location = new System.Drawing.Point(24, 520);
			this.ChkEmailCopyToSeller.Name = "ChkEmailCopyToSeller";
			this.ChkEmailCopyToSeller.Size = new System.Drawing.Size(136, 24);
			this.ChkEmailCopyToSeller.TabIndex = 69;
			this.ChkEmailCopyToSeller.Text = "Email Copy To Seller";
			// 
			// TxtShippingInsuranceCost
			// 
			this.TxtShippingInsuranceCost.Location = new System.Drawing.Point(200, 24);
			this.TxtShippingInsuranceCost.Name = "TxtShippingInsuranceCost";
			this.TxtShippingInsuranceCost.Size = new System.Drawing.Size(136, 20);
			this.TxtShippingInsuranceCost.TabIndex = 70;
			this.TxtShippingInsuranceCost.Text = "0";
			// 
			// LblShippingInsuranceCost
			// 
			this.LblShippingInsuranceCost.Location = new System.Drawing.Point(16, 24);
			this.LblShippingInsuranceCost.Name = "LblShippingInsuranceCost";
			this.LblShippingInsuranceCost.Size = new System.Drawing.Size(136, 23);
			this.LblShippingInsuranceCost.TabIndex = 71;
			this.LblShippingInsuranceCost.Text = "Shipping Insurance Cost:";
			// 
			// CboShippingService
			// 
			this.CboShippingService.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboShippingService.Location = new System.Drawing.Point(200, 48);
			this.CboShippingService.Name = "CboShippingService";
			this.CboShippingService.Size = new System.Drawing.Size(136, 21);
			this.CboShippingService.TabIndex = 73;
			// 
			// LblShippingService
			// 
			this.LblShippingService.Location = new System.Drawing.Point(16, 48);
			this.LblShippingService.Name = "LblShippingService";
			this.LblShippingService.Size = new System.Drawing.Size(96, 18);
			this.LblShippingService.TabIndex = 72;
			this.LblShippingService.Text = "Shipping Service:";
			// 
			// TxtShippingServiceCost
			// 
			this.TxtShippingServiceCost.Location = new System.Drawing.Point(200, 80);
			this.TxtShippingServiceCost.Name = "TxtShippingServiceCost";
			this.TxtShippingServiceCost.Size = new System.Drawing.Size(136, 20);
			this.TxtShippingServiceCost.TabIndex = 74;
			this.TxtShippingServiceCost.Text = "3.14";
			// 
			// LblShippingServiceCost
			// 
			this.LblShippingServiceCost.Location = new System.Drawing.Point(16, 80);
			this.LblShippingServiceCost.Name = "LblShippingServiceCost";
			this.LblShippingServiceCost.Size = new System.Drawing.Size(136, 23);
			this.LblShippingServiceCost.TabIndex = 75;
			this.LblShippingServiceCost.Text = "Shipping Service Cost:";
			// 
			// TxtShippingServiceAdditionalCost
			// 
			this.TxtShippingServiceAdditionalCost.Location = new System.Drawing.Point(200, 112);
			this.TxtShippingServiceAdditionalCost.Name = "TxtShippingServiceAdditionalCost";
			this.TxtShippingServiceAdditionalCost.Size = new System.Drawing.Size(136, 20);
			this.TxtShippingServiceAdditionalCost.TabIndex = 76;
			this.TxtShippingServiceAdditionalCost.Text = "2.14";
			// 
			// LblShippingServiceAdditionalCost
			// 
			this.LblShippingServiceAdditionalCost.Location = new System.Drawing.Point(16, 112);
			this.LblShippingServiceAdditionalCost.Name = "LblShippingServiceAdditionalCost";
			this.LblShippingServiceAdditionalCost.Size = new System.Drawing.Size(176, 23);
			this.LblShippingServiceAdditionalCost.TabIndex = 77;
			this.LblShippingServiceAdditionalCost.Text = "Shipping Service Additional Cost:";
			// 
			// TxtShippingServicePriority
			// 
			this.TxtShippingServicePriority.Location = new System.Drawing.Point(200, 144);
			this.TxtShippingServicePriority.Name = "TxtShippingServicePriority";
			this.TxtShippingServicePriority.Size = new System.Drawing.Size(136, 20);
			this.TxtShippingServicePriority.TabIndex = 78;
			this.TxtShippingServicePriority.Text = "1";
			// 
			// LblShippingServicePriority
			// 
			this.LblShippingServicePriority.Location = new System.Drawing.Point(16, 144);
			this.LblShippingServicePriority.Name = "LblShippingServicePriority";
			this.LblShippingServicePriority.Size = new System.Drawing.Size(136, 23);
			this.LblShippingServicePriority.TabIndex = 79;
			this.LblShippingServicePriority.Text = "Shipping Service Priority:";
			// 
			// GrpShippingServiceOptions
			// 
			this.GrpShippingServiceOptions.Controls.AddRange(new System.Windows.Forms.Control[] {
																									this.TxtShippingServiceCost,
																									this.TxtShippingServicePriority,
																									this.LblShippingServiceCost,
																									this.CboShippingService,
																									this.LblShippingService,
																									this.LblShippingServicePriority,
																									this.TxtShippingServiceAdditionalCost,
																									this.LblShippingInsuranceCost,
																									this.LblShippingServiceAdditionalCost,
																									this.TxtShippingInsuranceCost});
			this.GrpShippingServiceOptions.Location = new System.Drawing.Point(24, 40);
			this.GrpShippingServiceOptions.Name = "GrpShippingServiceOptions";
			this.GrpShippingServiceOptions.Size = new System.Drawing.Size(376, 176);
			this.GrpShippingServiceOptions.TabIndex = 80;
			this.GrpShippingServiceOptions.TabStop = false;
			this.GrpShippingServiceOptions.Text = "Shipping Service Options";
			// 
			// GrpSaleTaxType
			// 
			this.GrpSaleTaxType.Controls.AddRange(new System.Windows.Forms.Control[] {
																						 this.TxtSalesTaxAmount,
																						 this.LblSalesTaxAmount,
																						 this.ChkShippingIncludedInTax,
																						 this.TxtSalesTaxState,
																						 this.LblSalesTaxState,
																						 this.TxtSalesTaxPercent,
																						 this.LblSalesTaxPercent});
			this.GrpSaleTaxType.Location = new System.Drawing.Point(24, 224);
			this.GrpSaleTaxType.Name = "GrpSaleTaxType";
			this.GrpSaleTaxType.Size = new System.Drawing.Size(376, 152);
			this.GrpSaleTaxType.TabIndex = 81;
			this.GrpSaleTaxType.TabStop = false;
			this.GrpSaleTaxType.Text = "Sale Tax";
			// 
			// TxtSalesTaxAmount
			// 
			this.TxtSalesTaxAmount.Location = new System.Drawing.Point(136, 120);
			this.TxtSalesTaxAmount.Name = "TxtSalesTaxAmount";
			this.TxtSalesTaxAmount.Size = new System.Drawing.Size(160, 20);
			this.TxtSalesTaxAmount.TabIndex = 74;
			this.TxtSalesTaxAmount.Text = "3.14";
			// 
			// LblSalesTaxAmount
			// 
			this.LblSalesTaxAmount.Location = new System.Drawing.Point(16, 120);
			this.LblSalesTaxAmount.Name = "LblSalesTaxAmount";
			this.LblSalesTaxAmount.Size = new System.Drawing.Size(120, 18);
			this.LblSalesTaxAmount.TabIndex = 75;
			this.LblSalesTaxAmount.Text = "Sales Tax Amount:";
			// 
			// ChkShippingIncludedInTax
			// 
			this.ChkShippingIncludedInTax.Checked = true;
			this.ChkShippingIncludedInTax.CheckState = System.Windows.Forms.CheckState.Checked;
			this.ChkShippingIncludedInTax.Location = new System.Drawing.Point(16, 88);
			this.ChkShippingIncludedInTax.Name = "ChkShippingIncludedInTax";
			this.ChkShippingIncludedInTax.Size = new System.Drawing.Size(160, 24);
			this.ChkShippingIncludedInTax.TabIndex = 73;
			this.ChkShippingIncludedInTax.Text = "Shipping Included In Tax";
			// 
			// TxtSalesTaxState
			// 
			this.TxtSalesTaxState.Location = new System.Drawing.Point(136, 56);
			this.TxtSalesTaxState.Name = "TxtSalesTaxState";
			this.TxtSalesTaxState.Size = new System.Drawing.Size(160, 20);
			this.TxtSalesTaxState.TabIndex = 71;
			this.TxtSalesTaxState.Text = "CA";
			// 
			// LblSalesTaxState
			// 
			this.LblSalesTaxState.Location = new System.Drawing.Point(16, 56);
			this.LblSalesTaxState.Name = "LblSalesTaxState";
			this.LblSalesTaxState.Size = new System.Drawing.Size(120, 18);
			this.LblSalesTaxState.TabIndex = 72;
			this.LblSalesTaxState.Text = "Sales Tax State:";
			// 
			// TxtSalesTaxPercent
			// 
			this.TxtSalesTaxPercent.Location = new System.Drawing.Point(136, 24);
			this.TxtSalesTaxPercent.Name = "TxtSalesTaxPercent";
			this.TxtSalesTaxPercent.Size = new System.Drawing.Size(160, 20);
			this.TxtSalesTaxPercent.TabIndex = 69;
			this.TxtSalesTaxPercent.Text = "8.25";
			// 
			// LblSalesTaxPercent
			// 
			this.LblSalesTaxPercent.Location = new System.Drawing.Point(16, 24);
			this.LblSalesTaxPercent.Name = "LblSalesTaxPercent";
			this.LblSalesTaxPercent.Size = new System.Drawing.Size(120, 18);
			this.LblSalesTaxPercent.TabIndex = 70;
			this.LblSalesTaxPercent.Text = "SalesTax Percent:";
			// 
			// FrmSendInvoice
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(424, 673);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.GrpSaleTaxType,
																		  this.GrpShippingServiceOptions,
																		  this.ChkEmailCopyToSeller,
																		  this.TxtCheckoutInstructions,
																		  this.LblCheckoutInstructions,
																		  this.TxtPayPalEmailAddress,
																		  this.LblPayPalEmailAddress,
																		  this.CboPaymentMethod,
																		  this.LblPaymentMethod,
																		  this.TxtInsuranceFee,
																		  this.lblInsuranceFee,
																		  this.TxtTransactionId,
																		  this.LblTransactionId,
																		  this.CboInsuranceOption,
																		  this.LblInsuranceOption,
																		  this.TxtItemId,
																		  this.LblItemId,
																		  this.GrpResult,
																		  this.BtnSendInvoice});
			this.Name = "FrmSendInvoice";
			this.Text = "SendInvoice";
			this.Load += new System.EventHandler(this.FrmSendInvoice_Load);
			this.GrpResult.ResumeLayout(false);
			this.GrpShippingServiceOptions.ResumeLayout(false);
			this.GrpSaleTaxType.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmSendInvoice_Load(object sender, System.EventArgs e)
		{
			TxtItemId.Text = mItemID;
			string[] shippingservice = Enum.GetNames(typeof(ShippingServiceCodeType));
			foreach (string rsn in shippingservice)
			{
				if (rsn != "CustomCode")
				{
					CboShippingService.Items.Add(rsn);
				}
			}
			CboShippingService.SelectedIndex = 0;

			string[] insuranceoption = Enum.GetNames(typeof(InsuranceOptionCodeType));
			foreach (string rsn in insuranceoption)
			{
				if (rsn != "CustomCode")
				{
					CboInsuranceOption.Items.Add(rsn);
				}
			}
			CboInsuranceOption.SelectedIndex = 0;

			string[] paymentmethod = Enum.GetNames(typeof(BuyerPaymentMethodCodeType));
			foreach (string rsn in paymentmethod)
			{
				if (rsn != "CustomCode")
				{
					CboPaymentMethod.Items.Add(rsn);
				}
			}
			CboPaymentMethod.SelectedIndex = 0;
		}

		private void BtnSendInvoice_Click(object sender, System.EventArgs e)
		{
			try
			{
				SendInvoiceCall apicall = new SendInvoiceCall(Context);

				TxtStatus.Text = "";
				string ItemID = TxtItemId.Text;
				string TransactionID = TxtTransactionId.Text;

				ShippingServiceOptionsType ShippingServiceOption = new ShippingServiceOptionsType();
				
				ShippingServiceOption.ShippingInsuranceCost = new AmountType();
				ShippingServiceOption.ShippingInsuranceCost.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
				ShippingServiceOption.ShippingInsuranceCost.Value = Convert.ToDouble(TxtShippingInsuranceCost.Text);
					
				ShippingServiceOption.ShippingService = CboShippingService.SelectedItem.ToString();

				ShippingServiceOption.ShippingServiceCost = new AmountType();
				ShippingServiceOption.ShippingServiceCost.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
				ShippingServiceOption.ShippingServiceCost.Value=Convert.ToDouble(TxtShippingServiceCost.Text);

				
				ShippingServiceOption.ShippingServiceAdditionalCost = new AmountType();

				ShippingServiceOption.ShippingServiceAdditionalCost.Value=Convert.ToDouble(TxtShippingServiceAdditionalCost.Text);
				ShippingServiceOption.ShippingServiceAdditionalCost.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
				ShippingServiceOption.ShippingServicePriority = Int32.Parse(TxtShippingServicePriority.Text);

				//SaleTaxType related
				apicall.SalesTax = new SalesTaxType();
				apicall.SalesTax.SalesTaxPercent = (float)Convert.ToDouble(TxtSalesTaxPercent.Text);
				apicall.SalesTax.SalesTaxState = TxtSalesTaxState.Text;
				apicall.SalesTax.ShippingIncludedInTax = ChkShippingIncludedInTax.Checked;

				apicall.SalesTax.SalesTaxAmount = new AmountType();
				apicall.SalesTax.SalesTaxAmount.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
				apicall.SalesTax.SalesTaxAmount.Value = Convert.ToDouble(TxtSalesTaxAmount.Text);
				
        
				apicall.InsuranceOption = (InsuranceOptionCodeType)Enum.Parse(typeof(InsuranceOptionCodeType), CboInsuranceOption.SelectedItem.ToString());
				
				apicall.InsuranceFee = new AmountType();
				apicall.InsuranceFee.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
				apicall.InsuranceFee.Value = Convert.ToDouble(TxtInsuranceFee.Text);

				apicall.PaymentMethodsList = new BuyerPaymentMethodCodeTypeCollection();
				apicall.PaymentMethodsList.Add((BuyerPaymentMethodCodeType)Enum.Parse(typeof(BuyerPaymentMethodCodeType), CboPaymentMethod.SelectedItem.ToString()));
				
				apicall.PayPalEmailAddress = TxtPayPalEmailAddress.Text;
				apicall.CheckoutInstructions = TxtCheckoutInstructions.Text;
				apicall.EmailCopyToSeller = ChkEmailCopyToSeller.Checked;

				ShippingServiceOptionsTypeCollection ShippingServiceOptionsList= new ShippingServiceOptionsTypeCollection();
				ShippingServiceOptionsList.Add(ShippingServiceOption);

				apicall.SendInvoice(ItemID, TransactionID, ShippingServiceOptionsList);
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}
}
