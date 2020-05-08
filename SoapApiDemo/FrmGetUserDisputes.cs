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
	/// Summary description for GetUserDisputesForm.
	/// </summary>
	public class FrmGetUserDisputes : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Button BtnGetUserDisputes;
		private System.Windows.Forms.ComboBox CboSort;
		private System.Windows.Forms.Label LblSort;
		private System.Windows.Forms.ComboBox CboFilter;
		private System.Windows.Forms.Label LblFilter;
		private System.Windows.Forms.Label LblDisputes;
		private System.Windows.Forms.ListView LstDisputes;
		private System.Windows.Forms.ColumnHeader ClmDisputeId;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.ColumnHeader ClmTransactionId;
		private System.Windows.Forms.ColumnHeader ClmState;
		private System.Windows.Forms.ColumnHeader ClmStatus;
		private System.Windows.Forms.ColumnHeader ClmOtherPartyName;
		private System.Windows.Forms.ColumnHeader ClmRole;
		private System.Windows.Forms.ColumnHeader ClmDisputeReason;
		private System.Windows.Forms.ColumnHeader ClmDisputeExplanation;
		private System.Windows.Forms.ColumnHeader ClmDisputeModifiedTime;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetUserDisputes()
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
			this.BtnGetUserDisputes = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblDisputes = new System.Windows.Forms.Label();
			this.LstDisputes = new System.Windows.Forms.ListView();
			this.ClmDisputeId = new System.Windows.Forms.ColumnHeader();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmTransactionId = new System.Windows.Forms.ColumnHeader();
			this.ClmState = new System.Windows.Forms.ColumnHeader();
			this.ClmStatus = new System.Windows.Forms.ColumnHeader();
			this.ClmOtherPartyName = new System.Windows.Forms.ColumnHeader();
			this.ClmRole = new System.Windows.Forms.ColumnHeader();
			this.ClmDisputeReason = new System.Windows.Forms.ColumnHeader();
			this.ClmDisputeExplanation = new System.Windows.Forms.ColumnHeader();
			this.ClmDisputeModifiedTime = new System.Windows.Forms.ColumnHeader();
			this.CboSort = new System.Windows.Forms.ComboBox();
			this.LblSort = new System.Windows.Forms.Label();
			this.CboFilter = new System.Windows.Forms.ComboBox();
			this.LblFilter = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetUserDisputes
			// 
			this.BtnGetUserDisputes.Location = new System.Drawing.Point(208, 72);
			this.BtnGetUserDisputes.Name = "BtnGetUserDisputes";
			this.BtnGetUserDisputes.Size = new System.Drawing.Size(176, 23);
			this.BtnGetUserDisputes.TabIndex = 9;
			this.BtnGetUserDisputes.Text = "GetUserDisputes";
			this.BtnGetUserDisputes.Click += new System.EventHandler(this.BtnGetUserDisputes_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.AddRange(new System.Windows.Forms.Control[] {
																					this.LblDisputes,
																					this.LstDisputes});
			this.GrpResult.Location = new System.Drawing.Point(8, 104);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(848, 288);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblDisputes
			// 
			this.LblDisputes.Location = new System.Drawing.Point(16, 24);
			this.LblDisputes.Name = "LblDisputes";
			this.LblDisputes.Size = new System.Drawing.Size(475, 23);
			this.LblDisputes.TabIndex = 12;
			this.LblDisputes.Text = "User Disputes:";
			// 
			// LstDisputes
			// 
			this.LstDisputes.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						  this.ClmDisputeId,
																						  this.ClmItemId,
																						  this.ClmTransactionId,
																						  this.ClmState,
																						  this.ClmStatus,
																						  this.ClmOtherPartyName,
																						  this.ClmRole,
																						  this.ClmDisputeReason,
																						  this.ClmDisputeExplanation,
																						  this.ClmDisputeModifiedTime});
			this.LstDisputes.GridLines = true;
			this.LstDisputes.Location = new System.Drawing.Point(16, 56);
			this.LstDisputes.Name = "LstDisputes";
			this.LstDisputes.Size = new System.Drawing.Size(824, 224);
			this.LstDisputes.TabIndex = 13;
			this.LstDisputes.View = System.Windows.Forms.View.Details;
			// 
			// ClmDisputeId
			// 
			this.ClmDisputeId.Text = "Dispute Id";
			this.ClmDisputeId.Width = 65;
			// 
			// ClmItemId
			// 
			this.ClmItemId.Text = "Item Id";
			this.ClmItemId.Width = 83;
			// 
			// ClmTransactionId
			// 
			this.ClmTransactionId.Text = "Transaction Id";
			this.ClmTransactionId.Width = 83;
			// 
			// ClmState
			// 
			this.ClmState.Text = "State";
			this.ClmState.Width = 50;
			// 
			// ClmStatus
			// 
			this.ClmStatus.Text = "Status";
			this.ClmStatus.Width = 80;
			// 
			// ClmOtherPartyName
			// 
			this.ClmOtherPartyName.Text = "OtherPartyName";
			this.ClmOtherPartyName.Width = 90;
			// 
			// ClmRole
			// 
			this.ClmRole.Text = "My Role";
			this.ClmRole.Width = 55;
			// 
			// ClmDisputeReason
			// 
			this.ClmDisputeReason.Text = "DisputeReason";
			this.ClmDisputeReason.Width = 90;
			// 
			// ClmDisputeExplanation
			// 
			this.ClmDisputeExplanation.Text = "DisputeExplanation";
			this.ClmDisputeExplanation.Width = 112;
			// 
			// ClmDisputeModifiedTime
			// 
			this.ClmDisputeModifiedTime.Text = "DisputeModifiedTime";
			this.ClmDisputeModifiedTime.Width = 114;
			// 
			// CboSort
			// 
			this.CboSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboSort.Location = new System.Drawing.Point(256, 16);
			this.CboSort.Name = "CboSort";
			this.CboSort.Size = new System.Drawing.Size(144, 21);
			this.CboSort.TabIndex = 57;
			// 
			// LblSort
			// 
			this.LblSort.Location = new System.Drawing.Point(176, 16);
			this.LblSort.Name = "LblSort";
			this.LblSort.Size = new System.Drawing.Size(80, 18);
			this.LblSort.TabIndex = 56;
			this.LblSort.Text = "Sort:";
			// 
			// CboFilter
			// 
			this.CboFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboFilter.Location = new System.Drawing.Point(256, 40);
			this.CboFilter.Name = "CboFilter";
			this.CboFilter.Size = new System.Drawing.Size(144, 21);
			this.CboFilter.TabIndex = 59;
			// 
			// LblFilter
			// 
			this.LblFilter.Location = new System.Drawing.Point(176, 40);
			this.LblFilter.Name = "LblFilter";
			this.LblFilter.Size = new System.Drawing.Size(80, 18);
			this.LblFilter.TabIndex = 58;
			this.LblFilter.Text = "Reason:";
			// 
			// FrmGetUserDisputes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(872, 397);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.CboFilter,
																		  this.LblFilter,
																		  this.CboSort,
																		  this.LblSort,
																		  this.GrpResult,
																		  this.BtnGetUserDisputes});
			this.Name = "FrmGetUserDisputes";
			this.Text = "GetUserDisputesForm";
			this.Load += new System.EventHandler(this.FrmGetUserDisputes_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetUserDisputes_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstDisputes.Items.Clear();
	
				GetUserDisputesCall apicall = new GetUserDisputesCall(Context);

				apicall.DisputeFilterType = (DisputeFilterTypeCodeType) Enum.Parse(typeof(DisputeFilterTypeCodeType), CboFilter.SelectedItem.ToString());
				apicall.DisputeSortType = (DisputeSortTypeCodeType) Enum.Parse(typeof(DisputeSortTypeCodeType), CboSort.SelectedItem.ToString());

				PaginationType page = new PaginationType();
				page.PageNumber = 1;
				DisputeTypeCollection disputes = apicall.GetUserDisputes(page);
				
				foreach (DisputeType dsp in disputes)
				{
					string[] listparams = new string[10];
					listparams[0] = dsp.DisputeID;
					listparams[1] = dsp.Item.ItemID;
					listparams[2] =  dsp.TransactionID;
					listparams[3] =  dsp.DisputeState.ToString();
					listparams[4] =  dsp.DisputeStatus.ToString();
					listparams[5] =  dsp.OtherPartyName;
					listparams[6] =  dsp.UserRole.ToString();
					listparams[7] =  dsp.DisputeReason.ToString();
					listparams[8] =  dsp.DisputeExplanation.ToString();
					listparams[9] =  dsp.DisputeModifiedTime.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					this.LstDisputes.Items.Add(vi);

				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void FrmGetUserDisputes_Load(object sender, System.EventArgs e)
		{
			string[] sorts = Enum.GetNames(typeof(DisputeSortTypeCodeType));
			foreach (string srt in sorts)
			{
				if (srt != "CustomCode")
				{
					CboSort.Items.Add(srt);
				}
			}
			CboSort.SelectedIndex = 0;
			
			string[] filters = Enum.GetNames(typeof(DisputeFilterTypeCodeType));
			foreach (string fltr in filters)
			{
				if (fltr != "CustomCode")
				{
					CboFilter.Items.Add(fltr);
				}
			}
			CboFilter.SelectedIndex = 0;

	

		}

	}
}
