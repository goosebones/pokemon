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
	public class FrmAddDispute : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.ComboBox CboReason;
		private System.Windows.Forms.Label LblReason;
		private System.Windows.Forms.Button BtnAddDispute;
		private System.Windows.Forms.TextBox TxtTransactionId;
		private System.Windows.Forms.Label LblTransactionId;
		private System.Windows.Forms.ComboBox CboExplanation;
		private System.Windows.Forms.Label LblExplanation;
		private System.Windows.Forms.Label LblDisputeId;
		private System.Windows.Forms.TextBox TxtDisputeId;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmAddDispute()
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
			this.BtnAddDispute = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblDisputeId = new System.Windows.Forms.Label();
			this.TxtDisputeId = new System.Windows.Forms.TextBox();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.CboReason = new System.Windows.Forms.ComboBox();
			this.LblReason = new System.Windows.Forms.Label();
			this.TxtTransactionId = new System.Windows.Forms.TextBox();
			this.LblTransactionId = new System.Windows.Forms.Label();
			this.CboExplanation = new System.Windows.Forms.ComboBox();
			this.LblExplanation = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnAddDispute
			// 
			this.BtnAddDispute.Location = new System.Drawing.Point(120, 112);
			this.BtnAddDispute.Name = "BtnAddDispute";
			this.BtnAddDispute.Size = new System.Drawing.Size(112, 23);
			this.BtnAddDispute.TabIndex = 56;
			this.BtnAddDispute.Text = "AddDispute";
			this.BtnAddDispute.Click += new System.EventHandler(this.BtnAddDispute_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblDisputeId);
			this.GrpResult.Controls.Add(this.TxtDisputeId);
			this.GrpResult.Location = new System.Drawing.Point(8, 144);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(352, 64);
			this.GrpResult.TabIndex = 25;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblDisputeId
			// 
			this.LblDisputeId.Location = new System.Drawing.Point(16, 24);
			this.LblDisputeId.Name = "LblDisputeId";
			this.LblDisputeId.Size = new System.Drawing.Size(80, 23);
			this.LblDisputeId.TabIndex = 42;
			this.LblDisputeId.Text = "Dispute Id:";
			// 
			// TxtDisputeId
			// 
			this.TxtDisputeId.Location = new System.Drawing.Point(96, 24);
			this.TxtDisputeId.Name = "TxtDisputeId";
			this.TxtDisputeId.ReadOnly = true;
			this.TxtDisputeId.Size = new System.Drawing.Size(104, 20);
			this.TxtDisputeId.TabIndex = 41;
			this.TxtDisputeId.Text = "";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(152, 8);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(80, 20);
			this.TxtItemId.TabIndex = 26;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(72, 8);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(80, 23);
			this.LblItemId.TabIndex = 27;
			this.LblItemId.Text = "Item Id:";
			// 
			// CboReason
			// 
			this.CboReason.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboReason.Location = new System.Drawing.Point(152, 56);
			this.CboReason.Name = "CboReason";
			this.CboReason.Size = new System.Drawing.Size(144, 21);
			this.CboReason.TabIndex = 55;
			// 
			// LblReason
			// 
			this.LblReason.Location = new System.Drawing.Point(72, 56);
			this.LblReason.Name = "LblReason";
			this.LblReason.Size = new System.Drawing.Size(80, 18);
			this.LblReason.TabIndex = 54;
			this.LblReason.Text = "Reason:";
			// 
			// TxtTransactionId
			// 
			this.TxtTransactionId.Location = new System.Drawing.Point(152, 32);
			this.TxtTransactionId.Name = "TxtTransactionId";
			this.TxtTransactionId.Size = new System.Drawing.Size(80, 20);
			this.TxtTransactionId.TabIndex = 57;
			this.TxtTransactionId.Text = "0";
			// 
			// LblTransactionId
			// 
			this.LblTransactionId.Location = new System.Drawing.Point(72, 32);
			this.LblTransactionId.Name = "LblTransactionId";
			this.LblTransactionId.Size = new System.Drawing.Size(80, 23);
			this.LblTransactionId.TabIndex = 58;
			this.LblTransactionId.Text = "Transaction Id:";
			// 
			// CboExplanation
			// 
			this.CboExplanation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboExplanation.Location = new System.Drawing.Point(152, 80);
			this.CboExplanation.Name = "CboExplanation";
			this.CboExplanation.Size = new System.Drawing.Size(144, 21);
			this.CboExplanation.TabIndex = 60;
			// 
			// LblExplanation
			// 
			this.LblExplanation.Location = new System.Drawing.Point(72, 80);
			this.LblExplanation.Name = "LblExplanation";
			this.LblExplanation.Size = new System.Drawing.Size(80, 18);
			this.LblExplanation.TabIndex = 59;
			this.LblExplanation.Text = "Explanation:";
			// 
			// FrmAddDispute
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 213);
			this.Controls.Add(this.CboExplanation);
			this.Controls.Add(this.LblExplanation);
			this.Controls.Add(this.TxtTransactionId);
			this.Controls.Add(this.TxtItemId);
			this.Controls.Add(this.LblTransactionId);
			this.Controls.Add(this.CboReason);
			this.Controls.Add(this.LblReason);
			this.Controls.Add(this.LblItemId);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnAddDispute);
			this.Name = "FrmAddDispute";
			this.Text = "AddDispute";
			this.Load += new System.EventHandler(this.FrmAddDispute_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void FrmAddDispute_Load(object sender, System.EventArgs e)
		{
			string[] reasons = Enum.GetNames(typeof(DisputeReasonCodeType));
			foreach (string rsn in reasons)
			{
				if (rsn != "CustomCode")
				{
					CboReason.Items.Add(rsn);
				}
			}
			CboReason.SelectedIndex = 0;


			string[] explanation = Enum.GetNames(typeof(DisputeExplanationCodeType));
			foreach (string exp in explanation)
			{
				if (exp != "CustomCode")
				{
					CboExplanation.Items.Add(exp);
				}
			}
			CboExplanation.SelectedIndex = 0;

		}
		
		
		private void BtnAddDispute_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtDisputeId.Text = "";

				AddDisputeCall apicall = new AddDisputeCall(Context);
				
			
				string DisputeId = apicall.AddDispute(TxtItemId.Text, TxtTransactionId.Text,  (DisputeReasonCodeType) Enum.Parse(typeof(DisputeReasonCodeType), CboReason.SelectedItem.ToString()), (DisputeExplanationCodeType) Enum.Parse(typeof(DisputeExplanationCodeType), CboExplanation.SelectedItem.ToString()));
	
				TxtDisputeId.Text = DisputeId;

			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

	
	}
}
