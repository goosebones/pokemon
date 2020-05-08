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
	public class FrmGetOrders : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Button BtnGetOrders;
		private System.Windows.Forms.CheckBox ChkStartTo;
		private System.Windows.Forms.CheckBox ChkStartFrom;
		private System.Windows.Forms.Label LblTimeRange;
		private System.Windows.Forms.ComboBox CboRole;
		private System.Windows.Forms.Label LblRole;
		private System.Windows.Forms.ComboBox CboStatus;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.DateTimePicker DatePickStartTo;
		private System.Windows.Forms.DateTimePicker DatePickStartFrom;
		private System.Windows.Forms.Label LblStartSep;
		private System.Windows.Forms.Label LblOrders;
		private System.Windows.Forms.ListView LstOrders;
		private System.Windows.Forms.ColumnHeader ClmOrderId;
		private System.Windows.Forms.ColumnHeader ClmStatus;
		private System.Windows.Forms.ColumnHeader ClmCreator;
		private System.Windows.Forms.ColumnHeader ClmSaved;
		private System.Windows.Forms.ColumnHeader ClmItems;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetOrders()
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
			this.BtnGetOrders = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblOrders = new System.Windows.Forms.Label();
			this.LstOrders = new System.Windows.Forms.ListView();
			this.ClmOrderId = new System.Windows.Forms.ColumnHeader();
			this.ClmStatus = new System.Windows.Forms.ColumnHeader();
			this.ClmCreator = new System.Windows.Forms.ColumnHeader();
			this.ClmSaved = new System.Windows.Forms.ColumnHeader();
			this.ClmItems = new System.Windows.Forms.ColumnHeader();
			this.ChkStartTo = new System.Windows.Forms.CheckBox();
			this.ChkStartFrom = new System.Windows.Forms.CheckBox();
			this.LblTimeRange = new System.Windows.Forms.Label();
			this.CboRole = new System.Windows.Forms.ComboBox();
			this.LblRole = new System.Windows.Forms.Label();
			this.CboStatus = new System.Windows.Forms.ComboBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.DatePickStartTo = new System.Windows.Forms.DateTimePicker();
			this.DatePickStartFrom = new System.Windows.Forms.DateTimePicker();
			this.LblStartSep = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetOrders
			// 
			this.BtnGetOrders.Location = new System.Drawing.Point(192, 96);
			this.BtnGetOrders.Name = "BtnGetOrders";
			this.BtnGetOrders.Size = new System.Drawing.Size(128, 23);
			this.BtnGetOrders.TabIndex = 9;
			this.BtnGetOrders.Text = "GetOrders";
			this.BtnGetOrders.Click += new System.EventHandler(this.BtnGetOrders_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblOrders,
																					this.LstOrders});
			this.GrpResult.Location = new System.Drawing.Point(8, 128);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(544, 232);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblOrders
			// 
			this.LblOrders.Location = new System.Drawing.Point(8, 24);
			this.LblOrders.Name = "LblOrders";
			this.LblOrders.Size = new System.Drawing.Size(112, 23);
			this.LblOrders.TabIndex = 12;
			this.LblOrders.Text = "Orders:";
			// 
			// LstOrders
			// 
			this.LstOrders.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.ClmOrderId,
																						this.ClmStatus,
																						this.ClmCreator,
																						this.ClmSaved,
																						this.ClmItems});
			this.LstOrders.GridLines = true;
			this.LstOrders.Location = new System.Drawing.Point(8, 56);
			this.LstOrders.Name = "LstOrders";
			this.LstOrders.Size = new System.Drawing.Size(520, 168);
			this.LstOrders.TabIndex = 13;
			this.LstOrders.View = System.Windows.Forms.View.Details;
			// 
			// ClmOrderId
			// 
			this.ClmOrderId.Text = "Order Id";
			this.ClmOrderId.Width = 95;
			// 
			// ClmStatus
			// 
			this.ClmStatus.Text = "Status";
			this.ClmStatus.Width = 75;
			// 
			// ClmCreator
			// 
			this.ClmCreator.Text = "Creator";
			this.ClmCreator.Width = 84;
			// 
			// ClmSaved
			// 
			this.ClmSaved.Text = "Amount Saved";
			this.ClmSaved.Width = 91;
			// 
			// ClmItems
			// 
			this.ClmItems.Text = "Items";
			this.ClmItems.Width = 162;
			// 
			// ChkStartTo
			// 
			this.ChkStartTo.Location = new System.Drawing.Point(280, 64);
			this.ChkStartTo.Name = "ChkStartTo";
			this.ChkStartTo.Size = new System.Drawing.Size(12, 24);
			this.ChkStartTo.TabIndex = 93;
			// 
			// ChkStartFrom
			// 
			this.ChkStartFrom.Location = new System.Drawing.Point(96, 64);
			this.ChkStartFrom.Name = "ChkStartFrom";
			this.ChkStartFrom.Size = new System.Drawing.Size(12, 24);
			this.ChkStartFrom.TabIndex = 92;
			// 
			// LblTimeRange
			// 
			this.LblTimeRange.Location = new System.Drawing.Point(16, 64);
			this.LblTimeRange.Name = "LblTimeRange";
			this.LblTimeRange.Size = new System.Drawing.Size(72, 16);
			this.LblTimeRange.TabIndex = 91;
			this.LblTimeRange.Text = "Time Filter:";
			// 
			// CboRole
			// 
			this.CboRole.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboRole.Location = new System.Drawing.Point(96, 40);
			this.CboRole.Name = "CboRole";
			this.CboRole.Size = new System.Drawing.Size(144, 21);
			this.CboRole.TabIndex = 89;
			// 
			// LblRole
			// 
			this.LblRole.Location = new System.Drawing.Point(16, 40);
			this.LblRole.Name = "LblRole";
			this.LblRole.Size = new System.Drawing.Size(80, 18);
			this.LblRole.TabIndex = 88;
			this.LblRole.Text = "Role:";
			// 
			// CboStatus
			// 
			this.CboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboStatus.Location = new System.Drawing.Point(96, 16);
			this.CboStatus.Name = "CboStatus";
			this.CboStatus.Size = new System.Drawing.Size(144, 21);
			this.CboStatus.TabIndex = 87;
			// 
			// LblStatus
			// 
			this.LblStatus.Location = new System.Drawing.Point(16, 16);
			this.LblStatus.Name = "LblStatus";
			this.LblStatus.Size = new System.Drawing.Size(80, 18);
			this.LblStatus.TabIndex = 86;
			this.LblStatus.Text = "Status:";
			// 
			// DatePickStartTo
			// 
			this.DatePickStartTo.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickStartTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickStartTo.Location = new System.Drawing.Point(296, 64);
			this.DatePickStartTo.Name = "DatePickStartTo";
			this.DatePickStartTo.Size = new System.Drawing.Size(136, 20);
			this.DatePickStartTo.TabIndex = 83;
			// 
			// DatePickStartFrom
			// 
			this.DatePickStartFrom.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.DatePickStartFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.DatePickStartFrom.Location = new System.Drawing.Point(112, 64);
			this.DatePickStartFrom.Name = "DatePickStartFrom";
			this.DatePickStartFrom.Size = new System.Drawing.Size(136, 20);
			this.DatePickStartFrom.TabIndex = 82;
			// 
			// LblStartSep
			// 
			this.LblStartSep.Location = new System.Drawing.Point(256, 64);
			this.LblStartSep.Name = "LblStartSep";
			this.LblStartSep.Size = new System.Drawing.Size(16, 23);
			this.LblStartSep.TabIndex = 81;
			this.LblStartSep.Text = " - ";
			// 
			// FrmGetOrders
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(560, 365);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.ChkStartTo,
																		  this.ChkStartFrom,
																		  this.LblTimeRange,
																		  this.CboRole,
																		  this.LblRole,
																		  this.CboStatus,
																		  this.LblStatus,
																		  this.DatePickStartTo,
																		  this.DatePickStartFrom,
																		  this.LblStartSep,
																		  this.GrpResult,
																		  this.BtnGetOrders});
			this.Name = "FrmGetOrders";
			this.Text = "GetOrders";
			this.Load += new System.EventHandler(this.FrmGetOrders_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetOrders_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstOrders.Items.Clear();
	
				GetOrdersCall apicall = new GetOrdersCall(Context);

				TimeFilter fltr = new TimeFilter();
				if (ChkStartFrom.Checked)
				{
					fltr.TimeFrom = DatePickStartFrom.Value;
				}

				if (ChkStartTo.Checked)
				{
					fltr.TimeTo = DatePickStartTo.Value;
				}

				OrderTypeCollection orders = apicall.GetOrders(fltr, (TradingRoleCodeType) Enum.Parse(typeof(TradingRoleCodeType), CboRole.SelectedItem.ToString()), (OrderStatusCodeType) Enum.Parse(typeof(OrderStatusCodeType), CboStatus.SelectedItem.ToString()));

				foreach (OrderType order in orders)
				{
					string[] listparams = new string[5];
					listparams[0] = order.OrderID;
					listparams[1] = order.OrderStatus.ToString();
					listparams[2] = order.CreatingUserRole.ToString();
					listparams[3] = order.AmountSaved.Value.ToString();
					string[] itemids = new string[order.TransactionArray.Count];
					int indx = 0;
					foreach (TransactionType trans in order.TransactionArray)
					{
						itemids[indx] = trans.Item.ItemID;
						indx ++;
					}
					listparams[4] = String.Join(", ", itemids);
		
					ListViewItem vi = new ListViewItem(listparams);
					LstOrders.Items.Add(vi);

				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void FrmGetOrders_Load(object sender, System.EventArgs e)
		{
			DateTime now = DateTime.Now;
			DatePickStartTo.Value = now;
			DatePickStartFrom.Value = now.AddDays(-5);

			
			string[] roles = Enum.GetNames(typeof(TradingRoleCodeType));
			foreach (string rl in roles)
			{
				if (rl != "CustomCode")
				{
					CboRole.Items.Add(rl);
				}
			}
			CboRole.SelectedIndex = 0;

			string[] statie = Enum.GetNames(typeof(OrderStatusCodeType));
			foreach (string stat in statie)
			{
				if (stat != "CustomCode")
				{
					CboStatus.Items.Add(stat);
				}
			}
			CboStatus.SelectedIndex = 0;
		}

	}
	
}
