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
	/// Summary description for GetItemShippingForm.
	/// </summary>
	public class FrmGetItemShipping : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnGetItemShipping;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblDestination;
		private System.Windows.Forms.TextBox TxtDestination;
		private System.Windows.Forms.ListView LstShipSvc;
		private System.Windows.Forms.ColumnHeader ClmService;
		private System.Windows.Forms.ColumnHeader ClmCost;
		private System.Windows.Forms.ColumnHeader ClmInsurance;
		private System.Windows.Forms.TextBox TxtShipZip;
		internal System.Windows.Forms.Label LblShipZip;
		internal System.Windows.Forms.Label LblHandlingCost;
		private System.Windows.Forms.TextBox TxtHandlingCost;
		private System.Windows.Forms.TextBox TxtShipType;
		internal System.Windows.Forms.Label LblShipType;
		internal System.Windows.Forms.Label LblPackage;
		private System.Windows.Forms.TextBox TxtPackage;
		internal System.Windows.Forms.Label LblWeight;
		private System.Windows.Forms.TextBox TxtWeight;
		private System.Windows.Forms.ColumnHeader ClmAddedCost;
		private System.Windows.Forms.Label LblCountryCode;
		private System.Windows.Forms.Label LblQuantity;
		private System.Windows.Forms.TextBox TxtQuantity;
		private System.Windows.Forms.ComboBox CboCountry;
		private System.Windows.Forms.ColumnHeader ClmShipLocation;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetItemShipping()
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
			this.LblItemId = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.BtnGetItemShipping = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblWeight = new System.Windows.Forms.Label();
			this.TxtWeight = new System.Windows.Forms.TextBox();
			this.LblPackage = new System.Windows.Forms.Label();
			this.TxtPackage = new System.Windows.Forms.TextBox();
			this.LstShipSvc = new System.Windows.Forms.ListView();
			this.ClmService = new System.Windows.Forms.ColumnHeader();
			this.ClmCost = new System.Windows.Forms.ColumnHeader();
			this.ClmInsurance = new System.Windows.Forms.ColumnHeader();
			this.ClmAddedCost = new System.Windows.Forms.ColumnHeader();
			this.ClmShipLocation = new System.Windows.Forms.ColumnHeader();
			this.TxtShipZip = new System.Windows.Forms.TextBox();
			this.LblShipZip = new System.Windows.Forms.Label();
			this.LblHandlingCost = new System.Windows.Forms.Label();
			this.TxtHandlingCost = new System.Windows.Forms.TextBox();
			this.TxtShipType = new System.Windows.Forms.TextBox();
			this.LblShipType = new System.Windows.Forms.Label();
			this.LblDestination = new System.Windows.Forms.Label();
			this.TxtDestination = new System.Windows.Forms.TextBox();
			this.LblCountryCode = new System.Windows.Forms.Label();
			this.LblQuantity = new System.Windows.Forms.Label();
			this.TxtQuantity = new System.Windows.Forms.TextBox();
			this.CboCountry = new System.Windows.Forms.ComboBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(16, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(88, 16);
			this.LblItemId.TabIndex = 40;
			this.LblItemId.Text = "Item Id:";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(104, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(128, 20);
			this.TxtItemId.TabIndex = 27;
			this.TxtItemId.Text = "";
			// 
			// BtnGetItemShipping
			// 
			this.BtnGetItemShipping.Location = new System.Drawing.Point(192, 80);
			this.BtnGetItemShipping.Name = "BtnGetItemShipping";
			this.BtnGetItemShipping.Size = new System.Drawing.Size(104, 23);
			this.BtnGetItemShipping.TabIndex = 28;
			this.BtnGetItemShipping.Text = "GetItemShipping";
			this.BtnGetItemShipping.Click += new System.EventHandler(this.BtnGetItemShipping_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblWeight,
																					this.TxtWeight,
																					this.LblPackage,
																					this.TxtPackage,
																					this.LstShipSvc,
																					this.TxtShipZip,
																					this.LblShipZip,
																					this.LblHandlingCost,
																					this.TxtHandlingCost,
																					this.TxtShipType,
																					this.LblShipType});
			this.GrpResult.Location = new System.Drawing.Point(8, 112);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(464, 264);
			this.GrpResult.TabIndex = 41;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblWeight
			// 
			this.LblWeight.Location = new System.Drawing.Point(8, 120);
			this.LblWeight.Name = "LblWeight";
			this.LblWeight.Size = new System.Drawing.Size(120, 23);
			this.LblWeight.TabIndex = 79;
			this.LblWeight.Text = "Weight:";
			// 
			// TxtWeight
			// 
			this.TxtWeight.Location = new System.Drawing.Point(128, 120);
			this.TxtWeight.Name = "TxtWeight";
			this.TxtWeight.ReadOnly = true;
			this.TxtWeight.Size = new System.Drawing.Size(120, 20);
			this.TxtWeight.TabIndex = 80;
			this.TxtWeight.Text = "";
			// 
			// LblPackage
			// 
			this.LblPackage.Location = new System.Drawing.Point(8, 96);
			this.LblPackage.Name = "LblPackage";
			this.LblPackage.Size = new System.Drawing.Size(120, 23);
			this.LblPackage.TabIndex = 77;
			this.LblPackage.Text = "Package:";
			// 
			// TxtPackage
			// 
			this.TxtPackage.Location = new System.Drawing.Point(128, 96);
			this.TxtPackage.Name = "TxtPackage";
			this.TxtPackage.ReadOnly = true;
			this.TxtPackage.Size = new System.Drawing.Size(120, 20);
			this.TxtPackage.TabIndex = 78;
			this.TxtPackage.Text = "";
			// 
			// LstShipSvc
			// 
			this.LstShipSvc.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						 this.ClmService,
																						 this.ClmCost,
																						 this.ClmInsurance,
																						 this.ClmAddedCost,
																						 this.ClmShipLocation});
			this.LstShipSvc.GridLines = true;
			this.LstShipSvc.Location = new System.Drawing.Point(8, 168);
			this.LstShipSvc.Name = "LstShipSvc";
			this.LstShipSvc.Size = new System.Drawing.Size(424, 80);
			this.LstShipSvc.TabIndex = 76;
			this.LstShipSvc.View = System.Windows.Forms.View.Details;
			// 
			// ClmService
			// 
			this.ClmService.Text = "Service";
			this.ClmService.Width = 129;
			// 
			// ClmCost
			// 
			this.ClmCost.Text = "Cost";
			this.ClmCost.Width = 61;
			// 
			// ClmInsurance
			// 
			this.ClmInsurance.Text = "Insurance";
			this.ClmInsurance.Width = 63;
			// 
			// ClmAddedCost
			// 
			this.ClmAddedCost.Text = "Additional Cost";
			this.ClmAddedCost.Width = 83;
			// 
			// ClmShipLocation
			// 
			this.ClmShipLocation.Text = "Locations";
			this.ClmShipLocation.Width = 79;
			// 
			// TxtShipZip
			// 
			this.TxtShipZip.Location = new System.Drawing.Point(128, 48);
			this.TxtShipZip.Name = "TxtShipZip";
			this.TxtShipZip.ReadOnly = true;
			this.TxtShipZip.Size = new System.Drawing.Size(120, 20);
			this.TxtShipZip.TabIndex = 73;
			this.TxtShipZip.Text = "";
			// 
			// LblShipZip
			// 
			this.LblShipZip.Location = new System.Drawing.Point(8, 48);
			this.LblShipZip.Name = "LblShipZip";
			this.LblShipZip.Size = new System.Drawing.Size(120, 23);
			this.LblShipZip.TabIndex = 60;
			this.LblShipZip.Text = "Ship From Zip:";
			// 
			// LblHandlingCost
			// 
			this.LblHandlingCost.Location = new System.Drawing.Point(8, 72);
			this.LblHandlingCost.Name = "LblHandlingCost";
			this.LblHandlingCost.Size = new System.Drawing.Size(120, 23);
			this.LblHandlingCost.TabIndex = 62;
			this.LblHandlingCost.Text = "Handling Cost:";
			// 
			// TxtHandlingCost
			// 
			this.TxtHandlingCost.Location = new System.Drawing.Point(128, 72);
			this.TxtHandlingCost.Name = "TxtHandlingCost";
			this.TxtHandlingCost.ReadOnly = true;
			this.TxtHandlingCost.Size = new System.Drawing.Size(120, 20);
			this.TxtHandlingCost.TabIndex = 69;
			this.TxtHandlingCost.Text = "";
			// 
			// TxtShipType
			// 
			this.TxtShipType.Location = new System.Drawing.Point(128, 24);
			this.TxtShipType.Name = "TxtShipType";
			this.TxtShipType.ReadOnly = true;
			this.TxtShipType.Size = new System.Drawing.Size(120, 20);
			this.TxtShipType.TabIndex = 70;
			this.TxtShipType.Text = "";
			// 
			// LblShipType
			// 
			this.LblShipType.Location = new System.Drawing.Point(8, 24);
			this.LblShipType.Name = "LblShipType";
			this.LblShipType.Size = new System.Drawing.Size(120, 23);
			this.LblShipType.TabIndex = 54;
			this.LblShipType.Text = "Ship Type:";
			// 
			// LblDestination
			// 
			this.LblDestination.Location = new System.Drawing.Point(16, 32);
			this.LblDestination.Name = "LblDestination";
			this.LblDestination.Size = new System.Drawing.Size(88, 16);
			this.LblDestination.TabIndex = 43;
			this.LblDestination.Text = "Destination Zip:";
			// 
			// TxtDestination
			// 
			this.TxtDestination.Location = new System.Drawing.Point(104, 32);
			this.TxtDestination.Name = "TxtDestination";
			this.TxtDestination.Size = new System.Drawing.Size(128, 20);
			this.TxtDestination.TabIndex = 42;
			this.TxtDestination.Text = "";
			// 
			// LblCountryCode
			// 
			this.LblCountryCode.Location = new System.Drawing.Point(248, 8);
			this.LblCountryCode.Name = "LblCountryCode";
			this.LblCountryCode.Size = new System.Drawing.Size(96, 16);
			this.LblCountryCode.TabIndex = 45;
			this.LblCountryCode.Text = "Country Code:";
			// 
			// LblQuantity
			// 
			this.LblQuantity.Location = new System.Drawing.Point(248, 32);
			this.LblQuantity.Name = "LblQuantity";
			this.LblQuantity.Size = new System.Drawing.Size(96, 16);
			this.LblQuantity.TabIndex = 47;
			this.LblQuantity.Text = "Quantity:";
			// 
			// TxtQuantity
			// 
			this.TxtQuantity.Location = new System.Drawing.Point(344, 32);
			this.TxtQuantity.Name = "TxtQuantity";
			this.TxtQuantity.Size = new System.Drawing.Size(128, 20);
			this.TxtQuantity.TabIndex = 46;
			this.TxtQuantity.Text = "";
			// 
			// CboCountry
			// 
			this.CboCountry.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboCountry.Location = new System.Drawing.Point(344, 8);
			this.CboCountry.Name = "CboCountry";
			this.CboCountry.Size = new System.Drawing.Size(128, 21);
			this.CboCountry.TabIndex = 56;
			// 
			// FrmGetItemShipping
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 389);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.CboCountry,
																		  this.LblQuantity,
																		  this.TxtQuantity,
																		  this.LblCountryCode,
																		  this.LblDestination,
																		  this.TxtDestination,
																		  this.TxtItemId,
																		  this.GrpResult,
																		  this.LblItemId,
																		  this.BtnGetItemShipping});
			this.Name = "FrmGetItemShipping";
			this.Text = "GetItemShipping";
			this.Load += new System.EventHandler(this.FrmGetItemShipping_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetItemShipping_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstShipSvc.Items.Clear();
				TxtShipZip.Text = "";
				TxtHandlingCost.Text = "";
				TxtShipType.Text = "";
				TxtPackage.Text = "";
				TxtWeight.Text = "";


				GetItemShippingCall apicall = new GetItemShippingCall(Context);

				if (TxtQuantity.Text != String.Empty)
					apicall.QuantitySold = Convert.ToInt32(TxtQuantity.Text);
				
				if (CboCountry.SelectedIndex != 0)
					apicall.DestinationCountryCode = (CountryCodeType) Enum.Parse(typeof(CountryCodeType), CboCountry.SelectedItem.ToString());


				ShippingDetailsType shipdetails = apicall.GetItemShipping(TxtItemId.Text, TxtDestination.Text);

				TxtShipType.Text = shipdetails.ShippingType.ToString();

				if (shipdetails.CalculatedShippingRate != null)
				{
					TxtShipZip.Text = shipdetails.CalculatedShippingRate.OriginatingPostalCode;
					TxtHandlingCost.Text = shipdetails.CalculatedShippingRate.PackagingHandlingCosts.Value.ToString();
					//TxtPackage.Text = shipdetails.CalculatedShippingRate.ShippingPackage.ToString();
					//TxtWeight.Text = shipdetails.CalculatedShippingRate.WeightMajor.Value.ToString() + " " + shipdetails.CalculatedShippingRate.WeightMajor.unit + " - " + shipdetails.CalculatedShippingRate.WeightMinor.Value.ToString() + " " + shipdetails.CalculatedShippingRate.WeightMinor.unit;
				}

				foreach (ShippingServiceOptionsType shipopt in shipdetails.ShippingServiceOptions)
				{
					string[] listparams = new string[5];
					listparams[0] = shipopt.ShippingService.ToString();
					if (shipopt.ShippingServiceCost  != null)
						listparams[1] = shipopt.ShippingServiceCost.Value.ToString();
					if (shipopt.ShippingInsuranceCost  != null)
						listparams[2] = shipopt.ShippingInsuranceCost.Value.ToString();
					if (shipopt.ShippingServiceAdditionalCost  != null)
						listparams[3] = shipopt.ShippingServiceAdditionalCost.Value.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					this.LstShipSvc.Items.Add(vi);

				}
				foreach (InternationalShippingServiceOptionsType shipopt in shipdetails.InternationalShippingServiceOption)
				{
					string[] listparams = new string[5];
					listparams[0] = shipopt.ShippingService.ToString();
					if (shipopt.ShippingServiceCost  != null)
						listparams[1] = shipopt.ShippingServiceCost.Value.ToString();
					if (shipopt.ShippingServiceAdditionalCost  != null)
						listparams[3] = shipopt.ShippingServiceAdditionalCost.Value.ToString();
					listparams[4] = String.Join(", ",shipopt.ShipToLocation.ToArray());

					ListViewItem vi = new ListViewItem(listparams);
					this.LstShipSvc.Items.Add(vi);

				}

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void FrmGetItemShipping_Load(object sender, System.EventArgs e)
		{
			CboCountry.Items.Add("NoChange");
			
			string[] countries = Enum.GetNames(typeof(CountryCodeType));
			foreach (string cntry in countries)
			{
				if (cntry != "CustomCode")
				{
					CboCountry.Items.Add(cntry);
				}
			}
			CboCountry.SelectedIndex = 0;
		}

	
	}
}
