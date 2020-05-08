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
	public class FrmAddOrder : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Button BtnAddOrder;
		private System.Windows.Forms.TextBox TxtItemIdOne;
		private System.Windows.Forms.Label LblItemIdOne;
		private System.Windows.Forms.ComboBox CboShipSvc;
		private System.Windows.Forms.Label LblShipSvc;
		private System.Windows.Forms.TextBox TxtTransactionIdOne;
		private System.Windows.Forms.Label LblTransactionIdOne;
		private System.Windows.Forms.TextBox TxtTransactionIdTwo;
		private System.Windows.Forms.Label LblTransactionIdTwo;
		private System.Windows.Forms.TextBox TxtItemIdTwo;
		private System.Windows.Forms.Label LblItemIdTwo;
		private System.Windows.Forms.TextBox TxtShipCost;
		private System.Windows.Forms.Label LblShipCost;
		private System.Windows.Forms.TextBox TxtPaymentInstructions;
		private System.Windows.Forms.Label LblPaymentInstructions;
		private System.Windows.Forms.Label LblOrderId;
		private System.Windows.Forms.TextBox TxtOrderId;
		private System.Windows.Forms.TextBox TxtTotal;
		private System.Windows.Forms.Label LblTotal;
		private System.Windows.Forms.ComboBox CboRole;
		private System.Windows.Forms.Label LblUserRole;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmAddOrder()
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
			this.BtnAddOrder = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblOrderId = new System.Windows.Forms.Label();
			this.TxtOrderId = new System.Windows.Forms.TextBox();
			this.TxtItemIdOne = new System.Windows.Forms.TextBox();
			this.LblItemIdOne = new System.Windows.Forms.Label();
			this.CboShipSvc = new System.Windows.Forms.ComboBox();
			this.LblShipSvc = new System.Windows.Forms.Label();
			this.TxtTransactionIdOne = new System.Windows.Forms.TextBox();
			this.LblTransactionIdOne = new System.Windows.Forms.Label();
			this.TxtTransactionIdTwo = new System.Windows.Forms.TextBox();
			this.LblTransactionIdTwo = new System.Windows.Forms.Label();
			this.TxtItemIdTwo = new System.Windows.Forms.TextBox();
			this.LblItemIdTwo = new System.Windows.Forms.Label();
			this.TxtShipCost = new System.Windows.Forms.TextBox();
			this.LblShipCost = new System.Windows.Forms.Label();
			this.TxtPaymentInstructions = new System.Windows.Forms.TextBox();
			this.LblPaymentInstructions = new System.Windows.Forms.Label();
			this.TxtTotal = new System.Windows.Forms.TextBox();
			this.LblTotal = new System.Windows.Forms.Label();
			this.CboRole = new System.Windows.Forms.ComboBox();
			this.LblUserRole = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnAddOrder
			// 
			this.BtnAddOrder.Location = new System.Drawing.Point(120, 240);
			this.BtnAddOrder.Name = "BtnAddOrder";
			this.BtnAddOrder.Size = new System.Drawing.Size(112, 23);
			this.BtnAddOrder.TabIndex = 56;
			this.BtnAddOrder.Text = "AddOrder";
			this.BtnAddOrder.Click += new System.EventHandler(this.BtnAddOrder_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblOrderId,
																					this.TxtOrderId});
			this.GrpResult.Location = new System.Drawing.Point(16, 272);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(352, 64);
			this.GrpResult.TabIndex = 25;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblOrderId
			// 
			this.LblOrderId.Location = new System.Drawing.Point(16, 24);
			this.LblOrderId.Name = "LblOrderId";
			this.LblOrderId.Size = new System.Drawing.Size(80, 23);
			this.LblOrderId.TabIndex = 42;
			this.LblOrderId.Text = "Order Id:";
			// 
			// TxtOrderId
			// 
			this.TxtOrderId.Location = new System.Drawing.Point(96, 24);
			this.TxtOrderId.Name = "TxtOrderId";
			this.TxtOrderId.ReadOnly = true;
			this.TxtOrderId.Size = new System.Drawing.Size(104, 20);
			this.TxtOrderId.TabIndex = 41;
			this.TxtOrderId.Text = "";
			// 
			// TxtItemIdOne
			// 
			this.TxtItemIdOne.Location = new System.Drawing.Point(96, 8);
			this.TxtItemIdOne.Name = "TxtItemIdOne";
			this.TxtItemIdOne.Size = new System.Drawing.Size(80, 20);
			this.TxtItemIdOne.TabIndex = 26;
			this.TxtItemIdOne.Text = "";
			// 
			// LblItemIdOne
			// 
			this.LblItemIdOne.Location = new System.Drawing.Point(16, 8);
			this.LblItemIdOne.Name = "LblItemIdOne";
			this.LblItemIdOne.Size = new System.Drawing.Size(80, 23);
			this.LblItemIdOne.TabIndex = 27;
			this.LblItemIdOne.Text = "Item Id:";
			// 
			// CboShipSvc
			// 
			this.CboShipSvc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboShipSvc.Location = new System.Drawing.Point(96, 80);
			this.CboShipSvc.Name = "CboShipSvc";
			this.CboShipSvc.Size = new System.Drawing.Size(208, 21);
			this.CboShipSvc.TabIndex = 55;
			this.CboShipSvc.SelectedIndexChanged += new System.EventHandler(this.CboShipSvc_SelectedIndexChanged);
			// 
			// LblShipSvc
			// 
			this.LblShipSvc.Location = new System.Drawing.Point(16, 80);
			this.LblShipSvc.Name = "LblShipSvc";
			this.LblShipSvc.Size = new System.Drawing.Size(80, 18);
			this.LblShipSvc.TabIndex = 54;
			this.LblShipSvc.Text = "Service:";
			// 
			// TxtTransactionIdOne
			// 
			this.TxtTransactionIdOne.Location = new System.Drawing.Point(272, 8);
			this.TxtTransactionIdOne.Name = "TxtTransactionIdOne";
			this.TxtTransactionIdOne.Size = new System.Drawing.Size(80, 20);
			this.TxtTransactionIdOne.TabIndex = 57;
			this.TxtTransactionIdOne.Text = "";
			// 
			// LblTransactionIdOne
			// 
			this.LblTransactionIdOne.Location = new System.Drawing.Point(192, 8);
			this.LblTransactionIdOne.Name = "LblTransactionIdOne";
			this.LblTransactionIdOne.Size = new System.Drawing.Size(80, 23);
			this.LblTransactionIdOne.TabIndex = 58;
			this.LblTransactionIdOne.Text = "Transaction Id:";
			// 
			// TxtTransactionIdTwo
			// 
			this.TxtTransactionIdTwo.Location = new System.Drawing.Point(272, 32);
			this.TxtTransactionIdTwo.Name = "TxtTransactionIdTwo";
			this.TxtTransactionIdTwo.Size = new System.Drawing.Size(80, 20);
			this.TxtTransactionIdTwo.TabIndex = 61;
			this.TxtTransactionIdTwo.Text = "";
			// 
			// LblTransactionIdTwo
			// 
			this.LblTransactionIdTwo.Location = new System.Drawing.Point(192, 32);
			this.LblTransactionIdTwo.Name = "LblTransactionIdTwo";
			this.LblTransactionIdTwo.Size = new System.Drawing.Size(80, 23);
			this.LblTransactionIdTwo.TabIndex = 62;
			this.LblTransactionIdTwo.Text = "Transaction Id:";
			// 
			// TxtItemIdTwo
			// 
			this.TxtItemIdTwo.Location = new System.Drawing.Point(96, 32);
			this.TxtItemIdTwo.Name = "TxtItemIdTwo";
			this.TxtItemIdTwo.Size = new System.Drawing.Size(80, 20);
			this.TxtItemIdTwo.TabIndex = 63;
			this.TxtItemIdTwo.Text = "";
			// 
			// LblItemIdTwo
			// 
			this.LblItemIdTwo.Location = new System.Drawing.Point(16, 32);
			this.LblItemIdTwo.Name = "LblItemIdTwo";
			this.LblItemIdTwo.Size = new System.Drawing.Size(80, 23);
			this.LblItemIdTwo.TabIndex = 64;
			this.LblItemIdTwo.Text = "Item Id:";
			// 
			// TxtShipCost
			// 
			this.TxtShipCost.Location = new System.Drawing.Point(96, 104);
			this.TxtShipCost.Name = "TxtShipCost";
			this.TxtShipCost.Size = new System.Drawing.Size(80, 20);
			this.TxtShipCost.TabIndex = 65;
			this.TxtShipCost.Text = "";
			// 
			// LblShipCost
			// 
			this.LblShipCost.Location = new System.Drawing.Point(16, 104);
			this.LblShipCost.Name = "LblShipCost";
			this.LblShipCost.Size = new System.Drawing.Size(80, 23);
			this.LblShipCost.TabIndex = 66;
			this.LblShipCost.Text = "Shipping Cost:";
			// 
			// TxtPaymentInstructions
			// 
			this.TxtPaymentInstructions.Location = new System.Drawing.Point(96, 128);
			this.TxtPaymentInstructions.Multiline = true;
			this.TxtPaymentInstructions.Name = "TxtPaymentInstructions";
			this.TxtPaymentInstructions.Size = new System.Drawing.Size(256, 80);
			this.TxtPaymentInstructions.TabIndex = 67;
			this.TxtPaymentInstructions.Text = "";
			// 
			// LblPaymentInstructions
			// 
			this.LblPaymentInstructions.Location = new System.Drawing.Point(16, 128);
			this.LblPaymentInstructions.Name = "LblPaymentInstructions";
			this.LblPaymentInstructions.Size = new System.Drawing.Size(80, 23);
			this.LblPaymentInstructions.TabIndex = 68;
			this.LblPaymentInstructions.Text = "Instructions:";
			// 
			// TxtTotal
			// 
			this.TxtTotal.Location = new System.Drawing.Point(96, 208);
			this.TxtTotal.Name = "TxtTotal";
			this.TxtTotal.Size = new System.Drawing.Size(80, 20);
			this.TxtTotal.TabIndex = 69;
			this.TxtTotal.Text = "";
			// 
			// LblTotal
			// 
			this.LblTotal.Location = new System.Drawing.Point(16, 208);
			this.LblTotal.Name = "LblTotal";
			this.LblTotal.Size = new System.Drawing.Size(80, 23);
			this.LblTotal.TabIndex = 70;
			this.LblTotal.Text = "Total Amount:";
			// 
			// CboRole
			// 
			this.CboRole.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboRole.Location = new System.Drawing.Point(96, 56);
			this.CboRole.Name = "CboRole";
			this.CboRole.Size = new System.Drawing.Size(208, 21);
			this.CboRole.TabIndex = 72;
			// 
			// LblUserRole
			// 
			this.LblUserRole.Location = new System.Drawing.Point(16, 56);
			this.LblUserRole.Name = "LblUserRole";
			this.LblUserRole.Size = new System.Drawing.Size(80, 18);
			this.LblUserRole.TabIndex = 71;
			this.LblUserRole.Text = "Role:";
			// 
			// FrmAddOrder
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(384, 341);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.CboRole,
																		  this.LblUserRole,
																		  this.TxtTotal,
																		  this.LblTotal,
																		  this.TxtPaymentInstructions,
																		  this.LblPaymentInstructions,
																		  this.TxtShipCost,
																		  this.LblShipCost,
																		  this.TxtItemIdTwo,
																		  this.LblItemIdTwo,
																		  this.TxtTransactionIdTwo,
																		  this.LblTransactionIdTwo,
																		  this.TxtTransactionIdOne,
																		  this.TxtItemIdOne,
																		  this.LblTransactionIdOne,
																		  this.CboShipSvc,
																		  this.LblShipSvc,
																		  this.LblItemIdOne,
																		  this.GrpResult,
																		  this.BtnAddOrder});
			this.Name = "FrmAddOrder";
			this.Text = "AddOrder";
			this.Load += new System.EventHandler(this.FrmAddOrder_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmAddOrder_Load(object sender, System.EventArgs e)
		{
			string[] shpsvcs = Enum.GetNames(typeof(ShippingServiceCodeType));
			foreach (string svc in shpsvcs)
			{
				if (svc != "CustomCode")
				{
					CboShipSvc.Items.Add(svc);
				}
			}
			CboShipSvc.SelectedIndex = 0;

			string[] roles = Enum.GetNames(typeof(TradingRoleCodeType));
			foreach (string rol in roles)
			{
				if (rol != "CustomCode")
				{
					CboRole.Items.Add(rol);
				}
			}
			CboRole.SelectedIndex = 0;
		}
		
		
		private void BtnAddOrder_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtOrderId.Text = "";

				AddOrderCall apicall = new AddOrderCall(Context);
				

				OrderType order = new OrderType();
				order.TransactionArray = new TransactionTypeCollection();
				order.ShippingDetails = new ShippingDetailsType();
				order.PaymentMethods = new BuyerPaymentMethodCodeTypeCollection();
				
				TransactionType tr1 = new TransactionType();
				tr1.Item = new ItemType();
				tr1.Item.ItemID = TxtItemIdOne.Text;
				tr1.TransactionID = TxtTransactionIdOne.Text;
				order.TransactionArray.Add(tr1);
				
				TransactionType tr2 = new TransactionType();
				tr2.Item = new ItemType();
				tr2.Item.ItemID = TxtItemIdTwo.Text;
				tr2.TransactionID = TxtTransactionIdTwo.Text;
				order.TransactionArray.Add(tr2);

				order.ShippingDetails.PaymentInstructions = TxtPaymentInstructions.Text;
				ShippingServiceOptionsType shpopt = new ShippingServiceOptionsType();
				shpopt.ShippingService = CboShipSvc.SelectedItem.ToString();
				shpopt.ShippingServicePriority = 1;

				order.ShippingDetails.ShippingServiceOptions = new ShippingServiceOptionsTypeCollection();
				shpopt.ShippingServiceCost = new AmountType();
				shpopt.ShippingServiceCost.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
				
				if (TxtShipCost.Text.Length > 0) 
				{
					shpopt.ShippingServiceCost.Value = Convert.ToDouble(TxtShipCost.Text);
				}
				order.ShippingDetails.ShippingServiceOptions.Add(shpopt);
	
				order.Total = new AmountType();
				order.Total.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
				if (TxtTotal.Text.Length > 0)
					order.Total.Value = Convert.ToDouble(TxtTotal.Text);
				order.CreatingUserRole = (TradingRoleCodeType) Enum.Parse(typeof(TradingRoleCodeType), CboRole.SelectedItem.ToString());

				order.PaymentMethods.AddRange(new BuyerPaymentMethodCodeType[] {BuyerPaymentMethodCodeType.PaymentSeeDescription});
				
				string orderid = apicall.AddOrder(order);

	
				TxtOrderId.Text = orderid;

			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void CboShipSvc_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

	
	}
}
