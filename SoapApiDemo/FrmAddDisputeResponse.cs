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
	public class FrmAddDisputeResponse : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnAddDisputeResponse;
		private System.Windows.Forms.Label LblDisputeId;
		private System.Windows.Forms.TextBox TxtMessage;
		private System.Windows.Forms.Label LblMessage;
		private System.Windows.Forms.TextBox TxtDisputeId;
		private System.Windows.Forms.ComboBox CboActivity;
		private System.Windows.Forms.Label LblActivity;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.TextBox TxtShippingCarrier;
		private System.Windows.Forms.TextBox TxtShipmentTrackNumber;
		private System.Windows.Forms.Label LblShippingCarrier;
		private System.Windows.Forms.Label LblShipmentTrackNumber;
		private System.Windows.Forms.Label LblShippingTime;
		private System.Windows.Forms.DateTimePicker ShippingTimePick;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmAddDisputeResponse()
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
			this.BtnAddDisputeResponse = new System.Windows.Forms.Button();
			this.LblDisputeId = new System.Windows.Forms.Label();
			this.CboActivity = new System.Windows.Forms.ComboBox();
			this.LblActivity = new System.Windows.Forms.Label();
			this.TxtMessage = new System.Windows.Forms.TextBox();
			this.LblMessage = new System.Windows.Forms.Label();
			this.TxtDisputeId = new System.Windows.Forms.TextBox();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.TxtShippingCarrier = new System.Windows.Forms.TextBox();
			this.TxtShipmentTrackNumber = new System.Windows.Forms.TextBox();
			this.LblShippingCarrier = new System.Windows.Forms.Label();
			this.LblShipmentTrackNumber = new System.Windows.Forms.Label();
			this.LblShippingTime = new System.Windows.Forms.Label();
			this.ShippingTimePick = new System.Windows.Forms.DateTimePicker();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnAddDisputeResponse
			// 
			this.BtnAddDisputeResponse.Location = new System.Drawing.Point(128, 224);
			this.BtnAddDisputeResponse.Name = "BtnAddDisputeResponse";
			this.BtnAddDisputeResponse.Size = new System.Drawing.Size(136, 23);
			this.BtnAddDisputeResponse.TabIndex = 56;
			this.BtnAddDisputeResponse.Text = "AddDisputeResponse";
			this.BtnAddDisputeResponse.Click += new System.EventHandler(this.BtnAddDisputeResponse_Click);
			// 
			// LblDisputeId
			// 
			this.LblDisputeId.Location = new System.Drawing.Point(8, 8);
			this.LblDisputeId.Name = "LblDisputeId";
			this.LblDisputeId.Size = new System.Drawing.Size(80, 16);
			this.LblDisputeId.TabIndex = 27;
			this.LblDisputeId.Text = "Dispute Id:";
			// 
			// CboActivity
			// 
			this.CboActivity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboActivity.Location = new System.Drawing.Point(152, 96);
			this.CboActivity.Name = "CboActivity";
			this.CboActivity.Size = new System.Drawing.Size(152, 21);
			this.CboActivity.TabIndex = 55;
			this.CboActivity.SelectedIndexChanged += new System.EventHandler(this.CboActivity_SelectedIndexChanged);
			// 
			// LblActivity
			// 
			this.LblActivity.Location = new System.Drawing.Point(8, 96);
			this.LblActivity.Name = "LblActivity";
			this.LblActivity.Size = new System.Drawing.Size(80, 16);
			this.LblActivity.TabIndex = 54;
			this.LblActivity.Text = "Activity:";
			// 
			// TxtMessage
			// 
			this.TxtMessage.Location = new System.Drawing.Point(152, 32);
			this.TxtMessage.Multiline = true;
			this.TxtMessage.Name = "TxtMessage";
			this.TxtMessage.Size = new System.Drawing.Size(200, 56);
			this.TxtMessage.TabIndex = 57;
			this.TxtMessage.Text = "";
			// 
			// LblMessage
			// 
			this.LblMessage.Location = new System.Drawing.Point(8, 32);
			this.LblMessage.Name = "LblMessage";
			this.LblMessage.Size = new System.Drawing.Size(80, 16);
			this.LblMessage.TabIndex = 58;
			this.LblMessage.Text = "Message:";
			// 
			// TxtDisputeId
			// 
			this.TxtDisputeId.Location = new System.Drawing.Point(152, 8);
			this.TxtDisputeId.Name = "TxtDisputeId";
			this.TxtDisputeId.Size = new System.Drawing.Size(80, 20);
			this.TxtDisputeId.TabIndex = 26;
			this.TxtDisputeId.Text = "";
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblStatus,
																					this.TxtStatus});
			this.GrpResult.Location = new System.Drawing.Point(8, 256);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(352, 64);
			this.GrpResult.TabIndex = 59;
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
			// TxtShippingCarrier
			// 
			this.TxtShippingCarrier.Enabled = false;
			this.TxtShippingCarrier.Location = new System.Drawing.Point(152, 128);
			this.TxtShippingCarrier.Name = "TxtShippingCarrier";
			this.TxtShippingCarrier.Size = new System.Drawing.Size(152, 20);
			this.TxtShippingCarrier.TabIndex = 60;
			this.TxtShippingCarrier.Text = "";
			// 
			// TxtShipmentTrackNumber
			// 
			this.TxtShipmentTrackNumber.Enabled = false;
			this.TxtShipmentTrackNumber.Location = new System.Drawing.Point(152, 160);
			this.TxtShipmentTrackNumber.Name = "TxtShipmentTrackNumber";
			this.TxtShipmentTrackNumber.Size = new System.Drawing.Size(152, 20);
			this.TxtShipmentTrackNumber.TabIndex = 61;
			this.TxtShipmentTrackNumber.Text = "";
			// 
			// LblShippingCarrier
			// 
			this.LblShippingCarrier.Location = new System.Drawing.Point(8, 128);
			this.LblShippingCarrier.Name = "LblShippingCarrier";
			this.LblShippingCarrier.Size = new System.Drawing.Size(120, 16);
			this.LblShippingCarrier.TabIndex = 63;
			this.LblShippingCarrier.Text = "Shipping Carrier Used:";
			// 
			// LblShipmentTrackNumber
			// 
			this.LblShipmentTrackNumber.Location = new System.Drawing.Point(8, 160);
			this.LblShipmentTrackNumber.Name = "LblShipmentTrackNumber";
			this.LblShipmentTrackNumber.Size = new System.Drawing.Size(136, 16);
			this.LblShipmentTrackNumber.TabIndex = 64;
			this.LblShipmentTrackNumber.Text = "Shipment Track Number:";
			// 
			// LblShippingTime
			// 
			this.LblShippingTime.Location = new System.Drawing.Point(8, 192);
			this.LblShippingTime.Name = "LblShippingTime";
			this.LblShippingTime.Size = new System.Drawing.Size(100, 16);
			this.LblShippingTime.TabIndex = 65;
			this.LblShippingTime.Text = "Shipping Time:";
			// 
			// ShippingTimePick
			// 
			this.ShippingTimePick.CustomFormat = "yyyy-MM-dd HH:mm:ss";
			this.ShippingTimePick.Enabled = false;
			this.ShippingTimePick.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.ShippingTimePick.Location = new System.Drawing.Point(152, 192);
			this.ShippingTimePick.Name = "ShippingTimePick";
			this.ShippingTimePick.Size = new System.Drawing.Size(152, 20);
			this.ShippingTimePick.TabIndex = 66;
			// 
			// FrmAddDisputeResponse
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 345);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.ShippingTimePick,
																		  this.LblShippingTime,
																		  this.LblShipmentTrackNumber,
																		  this.LblShippingCarrier,
																		  this.TxtShipmentTrackNumber,
																		  this.TxtShippingCarrier,
																		  this.GrpResult,
																		  this.TxtMessage,
																		  this.TxtDisputeId,
																		  this.LblMessage,
																		  this.CboActivity,
																		  this.LblActivity,
																		  this.LblDisputeId,
																		  this.BtnAddDisputeResponse});
			this.Name = "FrmAddDisputeResponse";
			this.Text = "AddDisputeResponse";
			this.Load += new System.EventHandler(this.FrmAddDisputeResponse_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmAddDisputeResponse_Load(object sender, System.EventArgs e)
		{
			string[] activity = Enum.GetNames(typeof(DisputeActivityCodeType));
			foreach (string act in activity)
			{
				if (act != "CustomCode")
				{
					CboActivity.Items.Add(act);
				}
			}
			CboActivity.SelectedIndex = 0;



		}
		
		
		private void BtnAddDisputeResponse_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";

				AddDisputeResponseCall apicall = new AddDisputeResponseCall(Context);

				if (TxtShippingCarrier.Enabled)
				{
					apicall.ShippingCarrierUsed = TxtShippingCarrier.Text;
					apicall.ShipmentTrackNumber = TxtShipmentTrackNumber.Text;
					apicall.ShippingTime = ShippingTimePick.Value;
				}

				apicall.AddDisputeResponse(TxtDisputeId.Text, TxtMessage.Text,  (DisputeActivityCodeType) Enum.Parse(typeof(DisputeActivityCodeType), CboActivity.SelectedItem.ToString()));
	
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();

			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void CboActivity_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			string selectedText = CboActivity.Text;
			bool enabledFlag = selectedText.Equals("SellerShippedItem");
			
			TxtShippingCarrier.Enabled = enabledFlag;
			TxtShipmentTrackNumber.Enabled = enabledFlag;
			ShippingTimePick.Enabled = enabledFlag;
			
		}
	}
}
