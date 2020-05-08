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
	/// Summary description for GetApiAccessRulesForm.
	/// </summary>
	public class FrmGetApiAccessRules : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Button BtnGetApiAccessRules;
		private System.Windows.Forms.Label LblApiAccessRules;
		private System.Windows.Forms.ListView LstAccessRules;
		private System.Windows.Forms.ColumnHeader ClmCallName;
		private System.Windows.Forms.ColumnHeader ClmHourSoft;
		private System.Windows.Forms.ColumnHeader ClmHourHard;
		private System.Windows.Forms.ColumnHeader ClmHourUsage;
		private System.Windows.Forms.ColumnHeader ClmDaySoft;
		private System.Windows.Forms.ColumnHeader ClmDayHard;
		private System.Windows.Forms.ColumnHeader ClmDayUsage;
		private System.Windows.Forms.ColumnHeader ClmCounts;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetApiAccessRules()
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
			this.BtnGetApiAccessRules = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblApiAccessRules = new System.Windows.Forms.Label();
			this.LstAccessRules = new System.Windows.Forms.ListView();
			this.ClmCallName = new System.Windows.Forms.ColumnHeader();
			this.ClmHourSoft = new System.Windows.Forms.ColumnHeader();
			this.ClmHourHard = new System.Windows.Forms.ColumnHeader();
			this.ClmHourUsage = new System.Windows.Forms.ColumnHeader();
			this.ClmDaySoft = new System.Windows.Forms.ColumnHeader();
			this.ClmDayHard = new System.Windows.Forms.ColumnHeader();
			this.ClmDayUsage = new System.Windows.Forms.ColumnHeader();
			this.ClmCounts = new System.Windows.Forms.ColumnHeader();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetApiAccessRules
			// 
			this.BtnGetApiAccessRules.Location = new System.Drawing.Point(200, 8);
			this.BtnGetApiAccessRules.Name = "BtnGetApiAccessRules";
			this.BtnGetApiAccessRules.Size = new System.Drawing.Size(128, 23);
			this.BtnGetApiAccessRules.TabIndex = 9;
			this.BtnGetApiAccessRules.Text = "GetApiAccessRules";
			this.BtnGetApiAccessRules.Click += new System.EventHandler(this.BtnGetApiAccessRules_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblApiAccessRules);
			this.GrpResult.Controls.Add(this.LstAccessRules);
			this.GrpResult.Location = new System.Drawing.Point(8, 48);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(552, 280);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblApiAccessRules
			// 
			this.LblApiAccessRules.Location = new System.Drawing.Point(16, 24);
			this.LblApiAccessRules.Name = "LblApiAccessRules";
			this.LblApiAccessRules.Size = new System.Drawing.Size(475, 23);
			this.LblApiAccessRules.TabIndex = 12;
			this.LblApiAccessRules.Text = "Api Access Rules:";
			// 
			// LstAccessRules
			// 
			this.LstAccessRules.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																							 this.ClmCallName,
																							 this.ClmHourSoft,
																							 this.ClmHourHard,
																							 this.ClmHourUsage,
																							 this.ClmDaySoft,
																							 this.ClmDayHard,
																							 this.ClmDayUsage,
																							 this.ClmCounts});
			this.LstAccessRules.GridLines = true;
			this.LstAccessRules.Location = new System.Drawing.Point(16, 48);
			this.LstAccessRules.Name = "LstAccessRules";
			this.LstAccessRules.Size = new System.Drawing.Size(528, 224);
			this.LstAccessRules.TabIndex = 13;
			this.LstAccessRules.View = System.Windows.Forms.View.Details;
			// 
			// ClmCallName
			// 
			this.ClmCallName.Text = "Api Name";
			this.ClmCallName.Width = 69;
			// 
			// ClmHourSoft
			// 
			this.ClmHourSoft.Text = "Hour Soft";
			// 
			// ClmHourHard
			// 
			this.ClmHourHard.Text = "Hour Hard";
			this.ClmHourHard.Width = 64;
			// 
			// ClmHourUsage
			// 
			this.ClmHourUsage.Text = "Hour Usage";
			this.ClmHourUsage.Width = 73;
			// 
			// ClmDaySoft
			// 
			this.ClmDaySoft.Text = "Day Soft";
			this.ClmDaySoft.Width = 58;
			// 
			// ClmDayHard
			// 
			this.ClmDayHard.Text = "Day Hard";
			this.ClmDayHard.Width = 57;
			// 
			// ClmDayUsage
			// 
			this.ClmDayUsage.Text = "Day Usage";
			this.ClmDayUsage.Width = 66;
			// 
			// ClmCounts
			// 
			this.ClmCounts.Text = "Aggregates";
			this.ClmCounts.Width = 72;
			// 
			// FrmGetApiAccessRules
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(568, 341);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetApiAccessRules);
			this.Name = "FrmGetApiAccessRules";
			this.Text = "GetApiAccessRulesForm";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetApiAccessRules_Click(object sender, System.EventArgs e)
		{
			try
			{
				LstAccessRules.Items.Clear();
	
				GetApiAccessRulesCall apicall = new GetApiAccessRulesCall(Context);
				
				ApiAccessRuleTypeCollection rules = apicall.GetApiAccessRules();

				foreach (ApiAccessRuleType rule in rules)
				{
					string[] listparams = new string[8];
					listparams[0] = rule.CallName;
					listparams[1] = rule.HourlySoftLimit.ToString();
					listparams[2] = rule.HourlyHardLimit.ToString();
					listparams[3] = rule.HourlyUsage.ToString();
					listparams[4] = rule.DailySoftLimit.ToString();
					listparams[5] = rule.DailyHardLimit.ToString();
					listparams[6] = rule.DailyUsage.ToString();
					listparams[7] = rule.CountsTowardAggregate.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					this.LstAccessRules.Items.Add(vi);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}
	}
}
