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
	public class FrmSetStore : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.TextBox TxtDescription;
		private System.Windows.Forms.Label LblDescription;
		private System.Windows.Forms.Button BtnSetStore;
		private System.Windows.Forms.TextBox TxtName;
		private System.Windows.Forms.Label LblName;
		private System.Windows.Forms.TextBox TxtHeader;
		private System.Windows.Forms.Label LblHeader;
		private System.Windows.Forms.ComboBox CboLayout;
		private System.Windows.Forms.Label LblLayout;
		private System.Windows.Forms.ComboBox CboSort;
		private System.Windows.Forms.Label LblSort;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox TxtStatus;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmSetStore()
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
			this.TxtDescription = new System.Windows.Forms.TextBox();
			this.LblDescription = new System.Windows.Forms.Label();
			this.BtnSetStore = new System.Windows.Forms.Button();
			this.TxtName = new System.Windows.Forms.TextBox();
			this.LblName = new System.Windows.Forms.Label();
			this.TxtHeader = new System.Windows.Forms.TextBox();
			this.LblHeader = new System.Windows.Forms.Label();
			this.CboLayout = new System.Windows.Forms.ComboBox();
			this.LblLayout = new System.Windows.Forms.Label();
			this.CboSort = new System.Windows.Forms.ComboBox();
			this.LblSort = new System.Windows.Forms.Label();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.TxtStatus = new System.Windows.Forms.TextBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// TxtDescription
			// 
			this.TxtDescription.Location = new System.Drawing.Point(104, 40);
			this.TxtDescription.Multiline = true;
			this.TxtDescription.Name = "TxtDescription";
			this.TxtDescription.Size = new System.Drawing.Size(352, 56);
			this.TxtDescription.TabIndex = 35;
			this.TxtDescription.Text = "";
			// 
			// LblDescription
			// 
			this.LblDescription.Location = new System.Drawing.Point(24, 40);
			this.LblDescription.Name = "LblDescription";
			this.LblDescription.Size = new System.Drawing.Size(80, 18);
			this.LblDescription.TabIndex = 42;
			this.LblDescription.Text = "Description:";
			// 
			// BtnSetStore
			// 
			this.BtnSetStore.Location = new System.Drawing.Point(184, 208);
			this.BtnSetStore.Name = "BtnSetStore";
			this.BtnSetStore.Size = new System.Drawing.Size(120, 26);
			this.BtnSetStore.TabIndex = 36;
			this.BtnSetStore.Text = "SetStore";
			this.BtnSetStore.Click += new System.EventHandler(this.BtnSetStore_Click);
			// 
			// TxtName
			// 
			this.TxtName.Location = new System.Drawing.Point(104, 16);
			this.TxtName.Name = "TxtName";
			this.TxtName.Size = new System.Drawing.Size(248, 20);
			this.TxtName.TabIndex = 55;
			this.TxtName.Text = "";
			// 
			// LblName
			// 
			this.LblName.Location = new System.Drawing.Point(24, 16);
			this.LblName.Name = "LblName";
			this.LblName.Size = new System.Drawing.Size(80, 18);
			this.LblName.TabIndex = 56;
			this.LblName.Text = "Store Name:";
			// 
			// TxtHeader
			// 
			this.TxtHeader.Location = new System.Drawing.Point(104, 96);
			this.TxtHeader.Multiline = true;
			this.TxtHeader.Name = "TxtHeader";
			this.TxtHeader.Size = new System.Drawing.Size(352, 56);
			this.TxtHeader.TabIndex = 57;
			this.TxtHeader.Text = "";
			// 
			// LblHeader
			// 
			this.LblHeader.Location = new System.Drawing.Point(24, 96);
			this.LblHeader.Name = "LblHeader";
			this.LblHeader.Size = new System.Drawing.Size(80, 18);
			this.LblHeader.TabIndex = 58;
			this.LblHeader.Text = "Header:";
			// 
			// CboLayout
			// 
			this.CboLayout.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboLayout.Location = new System.Drawing.Point(104, 152);
			this.CboLayout.Name = "CboLayout";
			this.CboLayout.Size = new System.Drawing.Size(144, 21);
			this.CboLayout.TabIndex = 61;
			// 
			// LblLayout
			// 
			this.LblLayout.Location = new System.Drawing.Point(24, 152);
			this.LblLayout.Name = "LblLayout";
			this.LblLayout.Size = new System.Drawing.Size(80, 18);
			this.LblLayout.TabIndex = 60;
			this.LblLayout.Text = "Item Layout:";
			// 
			// CboSort
			// 
			this.CboSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboSort.Location = new System.Drawing.Point(104, 176);
			this.CboSort.Name = "CboSort";
			this.CboSort.Size = new System.Drawing.Size(144, 21);
			this.CboSort.TabIndex = 63;
			// 
			// LblSort
			// 
			this.LblSort.Location = new System.Drawing.Point(24, 176);
			this.LblSort.Name = "LblSort";
			this.LblSort.Size = new System.Drawing.Size(80, 18);
			this.LblSort.TabIndex = 62;
			this.LblSort.Text = "Item Sort:";
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.label1);
			this.GrpResult.Controls.Add(this.TxtStatus);
			this.GrpResult.Location = new System.Drawing.Point(16, 240);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(464, 64);
			this.GrpResult.TabIndex = 64;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 23);
			this.label1.TabIndex = 42;
			this.label1.Text = "Status:";
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
			// FrmSetStore
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 317);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.CboSort);
			this.Controls.Add(this.LblSort);
			this.Controls.Add(this.CboLayout);
			this.Controls.Add(this.LblLayout);
			this.Controls.Add(this.TxtHeader);
			this.Controls.Add(this.LblHeader);
			this.Controls.Add(this.TxtName);
			this.Controls.Add(this.LblName);
			this.Controls.Add(this.TxtDescription);
			this.Controls.Add(this.LblDescription);
			this.Controls.Add(this.BtnSetStore);
			this.Name = "FrmSetStore";
			this.Text = "SetStore";
			this.Load += new System.EventHandler(this.FrmSetStore_Load);
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
		private void FrmSetStore_Load(object sender, System.EventArgs e)
		{
			CboLayout.Items.Add("NoChange");
			
			string[] layouts = Enum.GetNames(typeof(StoreItemListLayoutCodeType));
			foreach (string layout in layouts)
			{
				if (layout != "CustomCode")
				{
					CboLayout.Items.Add(layout);
				}
			}
			CboLayout.SelectedIndex = 0;

			CboSort.Items.Add("NoChange");
			
			string[] sorts = Enum.GetNames(typeof(StoreItemListSortOrderCodeType));
			foreach (string srt in sorts)
			{
				if (srt != "CustomCode")
				{
					CboSort.Items.Add(srt);
				}
			}
			CboSort.SelectedIndex = 0;

		}


		private void BtnSetStore_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtStatus.Text = "";

				SetStoreCall apicall = new SetStoreCall(Context);

				StoreType store = new StoreType();
				
				if (TxtName.Text != String.Empty)
					store.Name = TxtName.Text;

				if (TxtDescription.Text != String.Empty)
					store.Description = TxtDescription.Text;

				if (TxtHeader.Text != String.Empty)
					store.CustomHeader = TxtHeader.Text;

				if (CboLayout.SelectedIndex > 0)
				{
					store.ItemListLayout = (StoreItemListLayoutCodeType) Enum.Parse(typeof(StoreItemListLayoutCodeType), CboLayout.SelectedItem.ToString());
				}

				if (CboSort.SelectedIndex > 0)
				{
					store.ItemListSortOrder = (StoreItemListSortOrderCodeType) Enum.Parse(typeof(StoreItemListSortOrderCodeType), CboSort.SelectedItem.ToString());
				}

				apicall.SetStore(store);
	
				TxtStatus.Text = apicall.ApiResponse.Ack.ToString();

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

	}
}
