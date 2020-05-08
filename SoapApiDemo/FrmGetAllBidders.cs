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
	public class FrmGetAllBidders : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.Label LblHighBidders;
		private System.Windows.Forms.Button BtnGetAllBidders;
		private System.Windows.Forms.ColumnHeader ClmAction;
		private System.Windows.Forms.ColumnHeader ClmUser;
		private System.Windows.Forms.ColumnHeader ClmCurrency;
		private System.Windows.Forms.ColumnHeader ClmMaxBid;
		private System.Windows.Forms.ColumnHeader ClmQuantiy;
		private System.Windows.Forms.ColumnHeader ClmTimeBid;
		private System.Windows.Forms.GroupBox GrpResults;
		private System.Windows.Forms.ListView LstHighBids;
		private System.Windows.Forms.ComboBox CboCallMode;
		private System.Windows.Forms.Label LblCallMode;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetAllBidders()
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
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.GrpResults = new System.Windows.Forms.GroupBox();
			this.LblHighBidders = new System.Windows.Forms.Label();
			this.LstHighBids = new System.Windows.Forms.ListView();
			this.ClmAction = new System.Windows.Forms.ColumnHeader();
			this.ClmUser = new System.Windows.Forms.ColumnHeader();
			this.ClmCurrency = new System.Windows.Forms.ColumnHeader();
			this.ClmMaxBid = new System.Windows.Forms.ColumnHeader();
			this.ClmQuantiy = new System.Windows.Forms.ColumnHeader();
			this.ClmTimeBid = new System.Windows.Forms.ColumnHeader();
			this.BtnGetAllBidders = new System.Windows.Forms.Button();
			this.CboCallMode = new System.Windows.Forms.ComboBox();
			this.LblCallMode = new System.Windows.Forms.Label();
			this.GrpResults.SuspendLayout();
			this.SuspendLayout();
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(256, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(88, 20);
			this.TxtItemId.TabIndex = 30;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(176, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(80, 18);
			this.LblItemId.TabIndex = 37;
			this.LblItemId.Text = "Item Id:";
			// 
			// GrpResults
			// 
			this.GrpResults.Controls.Add(this.LblHighBidders);
			this.GrpResults.Controls.Add(this.LstHighBids);
			this.GrpResults.Location = new System.Drawing.Point(8, 128);
			this.GrpResults.Name = "GrpResults";
			this.GrpResults.Size = new System.Drawing.Size(488, 192);
			this.GrpResults.TabIndex = 45;
			this.GrpResults.TabStop = false;
			this.GrpResults.Text = "Results";
			// 
			// LblHighBidders
			// 
			this.LblHighBidders.Location = new System.Drawing.Point(16, 24);
			this.LblHighBidders.Name = "LblHighBidders";
			this.LblHighBidders.Size = new System.Drawing.Size(136, 23);
			this.LblHighBidders.TabIndex = 14;
			this.LblHighBidders.Text = "High Bidders:";
			// 
			// LstHighBids
			// 
			this.LstHighBids.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						  this.ClmAction,
																						  this.ClmUser,
																						  this.ClmCurrency,
																						  this.ClmMaxBid,
																						  this.ClmQuantiy,
																						  this.ClmTimeBid});
			this.LstHighBids.GridLines = true;
			this.LstHighBids.Location = new System.Drawing.Point(24, 48);
			this.LstHighBids.Name = "LstHighBids";
			this.LstHighBids.Size = new System.Drawing.Size(456, 136);
			this.LstHighBids.TabIndex = 15;
			this.LstHighBids.View = System.Windows.Forms.View.Details;
			// 
			// ClmAction
			// 
			this.ClmAction.Text = "Action";
			this.ClmAction.Width = 50;
			// 
			// ClmUser
			// 
			this.ClmUser.Text = "User";
			this.ClmUser.Width = 59;
			// 
			// ClmCurrency
			// 
			this.ClmCurrency.Text = "Currency";
			this.ClmCurrency.Width = 62;
			// 
			// ClmMaxBid
			// 
			this.ClmMaxBid.Text = "MaxBid";
			// 
			// ClmQuantiy
			// 
			this.ClmQuantiy.Text = "Quantity";
			this.ClmQuantiy.Width = 59;
			// 
			// ClmTimeBid
			// 
			this.ClmTimeBid.Text = "Bid Time";
			this.ClmTimeBid.Width = 158;
			// 
			// BtnGetAllBidders
			// 
			this.BtnGetAllBidders.Location = new System.Drawing.Point(200, 88);
			this.BtnGetAllBidders.Name = "BtnGetAllBidders";
			this.BtnGetAllBidders.Size = new System.Drawing.Size(120, 26);
			this.BtnGetAllBidders.TabIndex = 46;
			this.BtnGetAllBidders.Text = "GetAllBidders";
			this.BtnGetAllBidders.Click += new System.EventHandler(this.BtnGetAllBidders_Click);
			// 
			// CboCallMode
			// 
			this.CboCallMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboCallMode.Location = new System.Drawing.Point(256, 32);
			this.CboCallMode.Name = "CboCallMode";
			this.CboCallMode.Size = new System.Drawing.Size(144, 21);
			this.CboCallMode.TabIndex = 57;
			// 
			// LblCallMode
			// 
			this.LblCallMode.Location = new System.Drawing.Point(176, 32);
			this.LblCallMode.Name = "LblCallMode";
			this.LblCallMode.Size = new System.Drawing.Size(80, 18);
			this.LblCallMode.TabIndex = 56;
			this.LblCallMode.Text = "Call Mode:";
			// 
			// FrmGetAllBidders
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 325);
			this.Controls.Add(this.CboCallMode);
			this.Controls.Add(this.LblCallMode);
			this.Controls.Add(this.BtnGetAllBidders);
			this.Controls.Add(this.GrpResults);
			this.Controls.Add(this.TxtItemId);
			this.Controls.Add(this.LblItemId);
			this.Name = "FrmGetAllBidders";
			this.Text = "GetAllBidders";
			this.Load += new System.EventHandler(this.FrmGetAllBidders_Load);
			this.GrpResults.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
	
		private void FrmGetAllBidders_Load(object sender, System.EventArgs e)
		{
			string[] modes = Enum.GetNames(typeof(GetAllBiddersModeCodeType));
			foreach (string mode in modes)
			{
				if (mode != "CustomCode")
				{
					CboCallMode.Items.Add(mode);
				}
			}
			CboCallMode.SelectedIndex = 0;

		
		}

		private void BtnGetAllBidders_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstHighBids.Items.Clear();
				
				GetAllBiddersCall apicall = new GetAllBiddersCall(Context);
				OfferTypeCollection	bids = apicall.GetAllBidders(TxtItemId.Text, (GetAllBiddersModeCodeType) Enum.Parse(typeof(GetAllBiddersModeCodeType), CboCallMode.SelectedItem.ToString()));

				foreach (OfferType offer in bids)
				{
					string[] listparams = new string[6];
					listparams[0] = offer.Action.ToString();
					listparams[1] = offer.User.UserID;
					listparams[2] = offer.Currency.ToString();
					listparams[3] = offer.MaxBid.Value.ToString();
					listparams[4] = offer.Quantity.ToString();
					listparams[5] = offer.TimeBid.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					this.LstHighBids.Items.Add(vi);

				}

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}




	}
}
