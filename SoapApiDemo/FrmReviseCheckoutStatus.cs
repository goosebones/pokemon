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
	public class FrmReviseCheckoutStatus : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.Button BtnReviseCheckoutStatus;
		private System.Windows.Forms.ComboBox CboCheckoutStatus;
		private System.Windows.Forms.Label LblCheckoutStatus;
		private System.Windows.Forms.ComboBox CboPaymentMethod;
		private System.Windows.Forms.Label LblPaymentMethod;
		private System.Windows.Forms.ComboBox CboShipSvc;
		private System.Windows.Forms.Label LblShipSvc;
		private System.Windows.Forms.TextBox TxtTransactionId;
		private System.Windows.Forms.Label LblTransactionId;
		private System.Windows.Forms.TextBox TxtOrderId;
		private System.Windows.Forms.Label LblOrderId;
		private System.Windows.Forms.TextBox TxtAmountPaid;
		private System.Windows.Forms.Label LblAmount;
		private System.Windows.Forms.RadioButton OptOrder;
		private System.Windows.Forms.RadioButton OptItem;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmReviseCheckoutStatus()
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
			this.BtnReviseCheckoutStatus = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.CboCheckoutStatus = new System.Windows.Forms.ComboBox();
			this.LblCheckoutStatus = new System.Windows.Forms.Label();
			this.CboPaymentMethod = new System.Windows.Forms.ComboBox();
			this.LblPaymentMethod = new System.Windows.Forms.Label();
			this.CboShipSvc = new System.Windows.Forms.ComboBox();
			this.LblShipSvc = new System.Windows.Forms.Label();
			this.TxtTransactionId = new System.Windows.Forms.TextBox();
			this.LblTransactionId = new System.Windows.Forms.Label();
			this.TxtOrderId = new System.Windows.Forms.TextBox();
			this.LblOrderId = new System.Windows.Forms.Label();
			this.TxtAmountPaid = new System.Windows.Forms.TextBox();
			this.LblAmount = new System.Windows.Forms.Label();
			this.OptOrder = new System.Windows.Forms.RadioButton();
			this.OptItem = new System.Windows.Forms.RadioButton();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnReviseCheckoutStatus
			// 
			this.BtnReviseCheckoutStatus.Location = new System.Drawing.Point(128, 168);
			this.BtnReviseCheckoutStatus.Name = "BtnReviseCheckoutStatus";
			this.BtnReviseCheckoutStatus.Size = new System.Drawing.Size(144, 23);
			this.BtnReviseCheckoutStatus.TabIndex = 23;
			this.BtnReviseCheckoutStatus.Text = "ReviseCheckoutStatus";
			this.BtnReviseCheckoutStatus.Click += new System.EventHandler(this.BtnReviseCheckoutStatus_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblStatus,
																					this.TxtStatus});
			this.GrpResult.Location = new System.Drawing.Point(32, 200);
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
			this.TxtItemId.Location = new System.Drawing.Point(160, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(72, 20);
			this.TxtItemId.TabIndex = 26;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(104, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(48, 23);
			this.LblItemId.TabIndex = 27;
			this.LblItemId.Text = "Item Id:";
			// 
			// CboCheckoutStatus
			// 
			this.CboCheckoutStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboCheckoutStatus.Location = new System.Drawing.Point(160, 88);
			this.CboCheckoutStatus.Name = "CboCheckoutStatus";
			this.CboCheckoutStatus.Size = new System.Drawing.Size(144, 21);
			this.CboCheckoutStatus.TabIndex = 55;
			// 
			// LblCheckoutStatus
			// 
			this.LblCheckoutStatus.Location = new System.Drawing.Point(48, 88);
			this.LblCheckoutStatus.Name = "LblCheckoutStatus";
			this.LblCheckoutStatus.Size = new System.Drawing.Size(112, 18);
			this.LblCheckoutStatus.TabIndex = 54;
			this.LblCheckoutStatus.Text = "Status:";
			// 
			// CboPaymentMethod
			// 
			this.CboPaymentMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboPaymentMethod.Location = new System.Drawing.Point(160, 112);
			this.CboPaymentMethod.Name = "CboPaymentMethod";
			this.CboPaymentMethod.Size = new System.Drawing.Size(144, 21);
			this.CboPaymentMethod.TabIndex = 57;
			// 
			// LblPaymentMethod
			// 
			this.LblPaymentMethod.Location = new System.Drawing.Point(48, 112);
			this.LblPaymentMethod.Name = "LblPaymentMethod";
			this.LblPaymentMethod.Size = new System.Drawing.Size(112, 18);
			this.LblPaymentMethod.TabIndex = 56;
			this.LblPaymentMethod.Text = "Payment Method:";
			// 
			// CboShipSvc
			// 
			this.CboShipSvc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboShipSvc.Location = new System.Drawing.Point(160, 136);
			this.CboShipSvc.Name = "CboShipSvc";
			this.CboShipSvc.Size = new System.Drawing.Size(144, 21);
			this.CboShipSvc.TabIndex = 59;
			// 
			// LblShipSvc
			// 
			this.LblShipSvc.Location = new System.Drawing.Point(48, 136);
			this.LblShipSvc.Name = "LblShipSvc";
			this.LblShipSvc.Size = new System.Drawing.Size(112, 18);
			this.LblShipSvc.TabIndex = 58;
			this.LblShipSvc.Text = "Shipping:";
			// 
			// TxtTransactionId
			// 
			this.TxtTransactionId.Location = new System.Drawing.Point(312, 8);
			this.TxtTransactionId.Name = "TxtTransactionId";
			this.TxtTransactionId.Size = new System.Drawing.Size(80, 20);
			this.TxtTransactionId.TabIndex = 60;
			this.TxtTransactionId.Text = "";
			// 
			// LblTransactionId
			// 
			this.LblTransactionId.Location = new System.Drawing.Point(232, 8);
			this.LblTransactionId.Name = "LblTransactionId";
			this.LblTransactionId.Size = new System.Drawing.Size(80, 23);
			this.LblTransactionId.TabIndex = 61;
			this.LblTransactionId.Text = "Transaction Id:";
			// 
			// TxtOrderId
			// 
			this.TxtOrderId.Location = new System.Drawing.Point(160, 32);
			this.TxtOrderId.Name = "TxtOrderId";
			this.TxtOrderId.Size = new System.Drawing.Size(72, 20);
			this.TxtOrderId.TabIndex = 62;
			this.TxtOrderId.Text = "";
			// 
			// LblOrderId
			// 
			this.LblOrderId.Location = new System.Drawing.Point(104, 32);
			this.LblOrderId.Name = "LblOrderId";
			this.LblOrderId.Size = new System.Drawing.Size(56, 23);
			this.LblOrderId.TabIndex = 63;
			this.LblOrderId.Text = "Order Id:";
			// 
			// TxtAmountPaid
			// 
			this.TxtAmountPaid.Location = new System.Drawing.Point(160, 64);
			this.TxtAmountPaid.Name = "TxtAmountPaid";
			this.TxtAmountPaid.Size = new System.Drawing.Size(80, 20);
			this.TxtAmountPaid.TabIndex = 64;
			this.TxtAmountPaid.Text = "";
			// 
			// LblAmount
			// 
			this.LblAmount.Location = new System.Drawing.Point(48, 64);
			this.LblAmount.Name = "LblAmount";
			this.LblAmount.Size = new System.Drawing.Size(112, 23);
			this.LblAmount.TabIndex = 65;
			this.LblAmount.Text = "Amount Paid:";
			// 
			// OptOrder
			// 
			this.OptOrder.Location = new System.Drawing.Point(8, 32);
			this.OptOrder.Name = "OptOrder";
			this.OptOrder.Size = new System.Drawing.Size(96, 24);
			this.OptOrder.TabIndex = 68;
			this.OptOrder.Text = "Revise Order:";
			this.OptOrder.CheckedChanged += new System.EventHandler(this.OptOrder_CheckedChanged);
			// 
			// OptItem
			// 
			this.OptItem.Location = new System.Drawing.Point(8, 8);
			this.OptItem.Name = "OptItem";
			this.OptItem.Size = new System.Drawing.Size(96, 24);
			this.OptItem.TabIndex = 67;
			this.OptItem.Text = "Revise Item:";
			this.OptItem.CheckedChanged += new System.EventHandler(this.OptItem_CheckedChanged);
			// 
			// FrmReviseCheckoutStatus
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(408, 277);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.OptOrder,
																		  this.OptItem,
																		  this.TxtAmountPaid,
																		  this.LblAmount,
																		  this.TxtOrderId,
																		  this.LblOrderId,
																		  this.TxtTransactionId,
																		  this.LblTransactionId,
																		  this.CboShipSvc,
																		  this.LblShipSvc,
																		  this.CboPaymentMethod,
																		  this.LblPaymentMethod,
																		  this.CboCheckoutStatus,
																		  this.LblCheckoutStatus,
																		  this.TxtItemId,
																		  this.LblItemId,
																		  this.GrpResult,
																		  this.BtnReviseCheckoutStatus});
			this.Name = "FrmReviseCheckoutStatus";
			this.Text = "ReviseCheckoutStatus";
			this.Load += new System.EventHandler(this.FrmReviseCheckoutStatus_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmReviseCheckoutStatus_Load(object sender, System.EventArgs e)
		{
			string[] paymethods = Enum.GetNames(typeof(BuyerPaymentMethodCodeType));
			foreach (string pay in paymethods)
			{
				if (pay != "CustomCode")
				{
					CboPaymentMethod.Items.Add(pay);
				}
			}

			CboPaymentMethod.SelectedIndex = 0;

			string[] svcs = Enum.GetNames(typeof(ShippingServiceCodeType));
			foreach (string svc in svcs)
			{
				if (svc != "CustomCode")
				{
					CboShipSvc.Items.Add(svc);
				}
			}
			CboShipSvc.SelectedIndex = 0;

			string[] statuses = Enum.GetNames(typeof(CompleteStatusCodeType));
			foreach (string stat in statuses)
			{
				if (stat != "CustomCode")
				{
					CboCheckoutStatus.Items.Add(stat);
				}
			}
			CboCheckoutStatus.SelectedIndex = 0;


			OptItem.Select();
		}
		
		
		private void BtnReviseCheckoutStatus_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";

				ReviseCheckoutStatusCall apicall = new ReviseCheckoutStatusCall(Context);
				
				apicall.PaymentMethodUsed = (BuyerPaymentMethodCodeType) Enum.Parse(typeof(BuyerPaymentMethodCodeType), CboPaymentMethod.SelectedItem.ToString());
				if (TxtAmountPaid.Text != String.Empty)
				{
					apicall.AmountPaid = new AmountType();
					apicall.AmountPaid.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
					apicall.AmountPaid.Value = Convert.ToDouble(TxtAmountPaid.Text);
				}

				apicall.ShippingService = CboShipSvc.SelectedItem.ToString();

				if (OptItem.Checked)
				{
					apicall.ReviseCheckoutStatus(TxtItemId.Text, TxtTransactionId.Text,(CompleteStatusCodeType) Enum.Parse(typeof(CompleteStatusCodeType), CboCheckoutStatus.SelectedItem.ToString()));


				} 
				else
				{
					apicall.ReviseCheckoutStatus(TxtOrderId.Text,(CompleteStatusCodeType) Enum.Parse(typeof(CompleteStatusCodeType), CboCheckoutStatus.SelectedItem.ToString()));

				}
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();

			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void OptItem_CheckedChanged(object sender, System.EventArgs e)
		{
			TxtItemId.Enabled = true;
			TxtTransactionId.Enabled = true;
			TxtOrderId.Enabled = false;

		}

		private void OptOrder_CheckedChanged(object sender, System.EventArgs e)
		{
			TxtItemId.Enabled = false;
			TxtTransactionId.Enabled = false;
			TxtOrderId.Enabled = true;
	
		}



	}
}
