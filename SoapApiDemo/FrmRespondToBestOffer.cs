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
	public class FrmRespondToBestOffer : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnRespondToBestOffer;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label LblStatus;
		private System.Windows.Forms.TextBox TxtStatus;
		private System.Windows.Forms.Label LblBestOfferID;
		private System.Windows.Forms.ComboBox CboAction;
		private System.Windows.Forms.Label LblAction;
		private System.Windows.Forms.Label LblSellerResponse;
		private System.Windows.Forms.TextBox TxtBestOfferID;
		private System.Windows.Forms.TextBox TxtItemID;
		private System.Windows.Forms.Label LblItemID;
		private System.Windows.Forms.TextBox TxtSellerResponse;
		private System.Windows.Forms.Label lblCounterOfferPrice;
		private System.Windows.Forms.TextBox txtCounterOfferPrice;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtCounterOfferQuantity;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmRespondToBestOffer()
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
			this.BtnRespondToBestOffer = new System.Windows.Forms.Button();
			this.LblBestOfferID = new System.Windows.Forms.Label();
			this.CboAction = new System.Windows.Forms.ComboBox();
			this.LblAction = new System.Windows.Forms.Label();
			this.TxtSellerResponse = new System.Windows.Forms.TextBox();
			this.LblSellerResponse = new System.Windows.Forms.Label();
			this.TxtBestOfferID = new System.Windows.Forms.TextBox();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblStatus = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.TxtItemID = new System.Windows.Forms.TextBox();
			this.LblItemID = new System.Windows.Forms.Label();
			this.txtCounterOfferPrice = new System.Windows.Forms.TextBox();
			this.lblCounterOfferPrice = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtCounterOfferQuantity = new System.Windows.Forms.TextBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnRespondToBestOffer
			// 
			this.BtnRespondToBestOffer.Location = new System.Drawing.Point(144, 288);
			this.BtnRespondToBestOffer.Name = "BtnRespondToBestOffer";
			this.BtnRespondToBestOffer.Size = new System.Drawing.Size(160, 23);
			this.BtnRespondToBestOffer.TabIndex = 56;
			this.BtnRespondToBestOffer.Text = "RespondToBestOffer";
			this.BtnRespondToBestOffer.Click += new System.EventHandler(this.BtnRespondToBestOffer_Click);
			// 
			// LblBestOfferID
			// 
			this.LblBestOfferID.Location = new System.Drawing.Point(16, 16);
			this.LblBestOfferID.Name = "LblBestOfferID";
			this.LblBestOfferID.Size = new System.Drawing.Size(80, 16);
			this.LblBestOfferID.TabIndex = 27;
			this.LblBestOfferID.Text = "BestOffer ID:";
			// 
			// CboAction
			// 
			this.CboAction.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboAction.Location = new System.Drawing.Point(160, 84);
			this.CboAction.Name = "CboAction";
			this.CboAction.Size = new System.Drawing.Size(152, 21);
			this.CboAction.TabIndex = 55;
			this.CboAction.SelectedIndexChanged += new System.EventHandler(this.CboAction_SelectedIndexChanged);
			// 
			// LblAction
			// 
			this.LblAction.Location = new System.Drawing.Point(16, 84);
			this.LblAction.Name = "LblAction";
			this.LblAction.Size = new System.Drawing.Size(80, 16);
			this.LblAction.TabIndex = 54;
			this.LblAction.Text = "Action:";
			// 
			// TxtSellerResponse
			// 
			this.TxtSellerResponse.Location = new System.Drawing.Point(160, 208);
			this.TxtSellerResponse.Multiline = true;
			this.TxtSellerResponse.Name = "TxtSellerResponse";
			this.TxtSellerResponse.Size = new System.Drawing.Size(224, 56);
			this.TxtSellerResponse.TabIndex = 57;
			this.TxtSellerResponse.Text = "";
			// 
			// LblSellerResponse
			// 
			this.LblSellerResponse.Location = new System.Drawing.Point(16, 208);
			this.LblSellerResponse.Name = "LblSellerResponse";
			this.LblSellerResponse.Size = new System.Drawing.Size(104, 16);
			this.LblSellerResponse.TabIndex = 58;
			this.LblSellerResponse.Text = "Seller Response:";
			// 
			// TxtBestOfferID
			// 
			this.TxtBestOfferID.Location = new System.Drawing.Point(160, 16);
			this.TxtBestOfferID.Name = "TxtBestOfferID";
			this.TxtBestOfferID.Size = new System.Drawing.Size(152, 20);
			this.TxtBestOfferID.TabIndex = 26;
			this.TxtBestOfferID.Text = "";
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblStatus,
																					this.TxtStatus});
			this.GrpResult.Location = new System.Drawing.Point(16, 336);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(304, 64);
			this.GrpResult.TabIndex = 59;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblStatus
			// 
			this.LblStatus.Location = new System.Drawing.Point(32, 24);
			this.LblStatus.Name = "LblStatus";
			this.LblStatus.Size = new System.Drawing.Size(80, 23);
			this.LblStatus.TabIndex = 42;
			this.LblStatus.Text = "Status:";
			// 
			// TxtStatus
			// 
			this.TxtStatus.Location = new System.Drawing.Point(112, 24);
			this.TxtStatus.Name = "TxtStatus";
			this.TxtStatus.ReadOnly = true;
			this.TxtStatus.Size = new System.Drawing.Size(72, 20);
			this.TxtStatus.TabIndex = 41;
			this.TxtStatus.Text = "";
			// 
			// TxtItemID
			// 
			this.TxtItemID.Location = new System.Drawing.Point(160, 50);
			this.TxtItemID.Name = "TxtItemID";
			this.TxtItemID.Size = new System.Drawing.Size(152, 20);
			this.TxtItemID.TabIndex = 60;
			this.TxtItemID.Text = "";
			// 
			// LblItemID
			// 
			this.LblItemID.Location = new System.Drawing.Point(16, 50);
			this.LblItemID.Name = "LblItemID";
			this.LblItemID.Size = new System.Drawing.Size(120, 16);
			this.LblItemID.TabIndex = 63;
			this.LblItemID.Text = "Item ID:";
			// 
			// txtCounterOfferPrice
			// 
			this.txtCounterOfferPrice.Location = new System.Drawing.Point(160, 128);
			this.txtCounterOfferPrice.Name = "txtCounterOfferPrice";
			this.txtCounterOfferPrice.Size = new System.Drawing.Size(152, 20);
			this.txtCounterOfferPrice.TabIndex = 64;
			this.txtCounterOfferPrice.Text = "";
			// 
			// lblCounterOfferPrice
			// 
			this.lblCounterOfferPrice.Location = new System.Drawing.Point(16, 120);
			this.lblCounterOfferPrice.Name = "lblCounterOfferPrice";
			this.lblCounterOfferPrice.Size = new System.Drawing.Size(104, 24);
			this.lblCounterOfferPrice.TabIndex = 65;
			this.lblCounterOfferPrice.Text = "Counter Offer Price:";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 160);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(104, 32);
			this.label1.TabIndex = 65;
			this.label1.Text = "Counter Offer Quantity:";
			// 
			// txtCounterOfferQuantity
			// 
			this.txtCounterOfferQuantity.Location = new System.Drawing.Point(160, 168);
			this.txtCounterOfferQuantity.Name = "txtCounterOfferQuantity";
			this.txtCounterOfferQuantity.Size = new System.Drawing.Size(152, 20);
			this.txtCounterOfferQuantity.TabIndex = 64;
			this.txtCounterOfferQuantity.Text = "";
			// 
			// FrmRespondToBestOffer
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(512, 446);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.lblCounterOfferPrice,
																		  this.txtCounterOfferPrice,
																		  this.LblItemID,
																		  this.TxtItemID,
																		  this.GrpResult,
																		  this.TxtSellerResponse,
																		  this.TxtBestOfferID,
																		  this.LblSellerResponse,
																		  this.CboAction,
																		  this.LblAction,
																		  this.LblBestOfferID,
																		  this.BtnRespondToBestOffer,
																		  this.label1,
																		  this.txtCounterOfferQuantity});
			this.Name = "FrmRespondToBestOffer";
			this.Text = "RespondToBestOffer";
			this.Load += new System.EventHandler(this.FrmRespondToBestOffer_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmRespondToBestOffer_Load(object sender, System.EventArgs e)
		{
			string[] action = Enum.GetNames(typeof(BestOfferActionCodeType));
			foreach (string act in action)
			{
				if (act != "CustomCode")
				{
					CboAction.Items.Add(act);
				}
			}
			CboAction.SelectedIndex = 0;



		}
		
		
		private void BtnRespondToBestOffer_Click(object sender, System.EventArgs e)
		{
			try
			{
				if(TxtBestOfferID.Text.Trim().Length == 0)
				{
					MessageBox.Show("Please input best offer id!");
					return;
				}

				if(TxtItemID.Text.Trim().Length == 0)
				{
					MessageBox.Show("Please input item id!");
					return;
				}


				TxtStatus.Text = "";
				RespondToBestOfferCall apicall = new RespondToBestOfferCall(Context);

				StringCollection BestOfferIDList = new StringCollection();
				BestOfferIDList.Add(TxtBestOfferID.Text);
				string cboActionChoice = CboAction.SelectedItem.ToString();
				if(cboActionChoice == "Counter") 
				{
						if(txtCounterOfferPrice.Text.Trim().Length == 0)
						{
							MessageBox.Show("Please counter offer price!");
							return;
						}

						if(txtCounterOfferQuantity.Text.Trim().Length == 0)
						{
							MessageBox.Show("Please input counter offer quanity!");
							return;
						}

						AmountType CounterOfferPrice = new AmountType();
						CounterOfferPrice.currencyID = CurrencyUtility.GetDefaultCurrencyCodeType(Context.Site);
						CounterOfferPrice.Value = Convert.ToDouble(txtCounterOfferPrice.Text);
					    int quantity = Convert.ToInt32(txtCounterOfferQuantity.Text);
						apicall.RespondToBestOffer(TxtItemID.Text, 
													BestOfferIDList,			
													(BestOfferActionCodeType) Enum.Parse(typeof(BestOfferActionCodeType), 
													CboAction.SelectedItem.ToString()), 
													TxtSellerResponse.Text,
													CounterOfferPrice,
													quantity);
				}
				else 
				{
					apicall.RespondToBestOffer(TxtItemID.Text, 
										   BestOfferIDList,			
										   (BestOfferActionCodeType) Enum.Parse(typeof(BestOfferActionCodeType), 
											CboAction.SelectedItem.ToString()), 
											TxtSellerResponse.Text);
				}
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();

			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void CboAction_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			string cboActionChoice = CboAction.SelectedItem.ToString();
			if(cboActionChoice == "Counter") 
			{
				txtCounterOfferPrice.Enabled = true;
				txtCounterOfferQuantity.Enabled = true;
			}
			else 
			{
				txtCounterOfferPrice.Enabled = false;
				txtCounterOfferQuantity.Enabled = false;
			}
		}
	}
}
