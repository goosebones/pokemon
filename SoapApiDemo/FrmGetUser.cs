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
	public class FrmGetUser : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Label LblUserIdRet;
		private System.Windows.Forms.TextBox TxtUserId;
		private System.Windows.Forms.Button BtnGetUser;
		private System.Windows.Forms.Label LblUserId;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.TextBox TxtStar;
		private System.Windows.Forms.Label LblStar;
		private System.Windows.Forms.TextBox TxtStoreUrl;
		private System.Windows.Forms.Label LblStoreUrl;
		private System.Windows.Forms.TextBox TxtSite;
		private System.Windows.Forms.Label LblSite;
		private System.Windows.Forms.TextBox TxtRegDate;
		private System.Windows.Forms.TextBox TxtEmail;
		private System.Windows.Forms.TextBox TxtFeedBackScore;
		private System.Windows.Forms.TextBox TxtSellerLevel;
		private System.Windows.Forms.Label LblEmail;
		private System.Windows.Forms.Label LblRegDate;
		private System.Windows.Forms.Label LblFeedBackScore;
		private System.Windows.Forms.Label LblSellerLevel;
		private System.Windows.Forms.TextBox TxtItemId;
		private System.Windows.Forms.Label LblItemId;
		private System.Windows.Forms.TextBox TxtUserIdRet;
		private System.Windows.Forms.Label LblVerified;
		private System.Windows.Forms.TextBox TxtVerified;
		private System.Windows.Forms.TextBox TxtChanged;
		private System.Windows.Forms.Label LblLastChanged;
		private System.Windows.Forms.TextBox TxtNew;
		private System.Windows.Forms.Label LblNew;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetUser()
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
			this.TxtUserId = new System.Windows.Forms.TextBox();
			this.BtnGetUser = new System.Windows.Forms.Button();
			this.LblUserId = new System.Windows.Forms.Label();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.TxtNew = new System.Windows.Forms.TextBox();
			this.LblNew = new System.Windows.Forms.Label();
			this.TxtChanged = new System.Windows.Forms.TextBox();
			this.LblLastChanged = new System.Windows.Forms.Label();
			this.TxtVerified = new System.Windows.Forms.TextBox();
			this.LblVerified = new System.Windows.Forms.Label();
			this.TxtUserIdRet = new System.Windows.Forms.TextBox();
			this.LblUserIdRet = new System.Windows.Forms.Label();
			this.TxtStar = new System.Windows.Forms.TextBox();
			this.LblStar = new System.Windows.Forms.Label();
			this.TxtStoreUrl = new System.Windows.Forms.TextBox();
			this.LblStoreUrl = new System.Windows.Forms.Label();
			this.TxtSite = new System.Windows.Forms.TextBox();
			this.LblSite = new System.Windows.Forms.Label();
			this.TxtRegDate = new System.Windows.Forms.TextBox();
			this.TxtEmail = new System.Windows.Forms.TextBox();
			this.TxtFeedBackScore = new System.Windows.Forms.TextBox();
			this.TxtSellerLevel = new System.Windows.Forms.TextBox();
			this.LblEmail = new System.Windows.Forms.Label();
			this.LblRegDate = new System.Windows.Forms.Label();
			this.LblFeedBackScore = new System.Windows.Forms.Label();
			this.LblSellerLevel = new System.Windows.Forms.Label();
			this.TxtItemId = new System.Windows.Forms.TextBox();
			this.LblItemId = new System.Windows.Forms.Label();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// TxtUserId
			// 
			this.TxtUserId.Location = new System.Drawing.Point(96, 16);
			this.TxtUserId.Name = "TxtUserId";
			this.TxtUserId.Size = new System.Drawing.Size(80, 20);
			this.TxtUserId.TabIndex = 22;
			this.TxtUserId.Text = "";
			// 
			// BtnGetUser
			// 
			this.BtnGetUser.Location = new System.Drawing.Point(136, 48);
			this.BtnGetUser.Name = "BtnGetUser";
			this.BtnGetUser.Size = new System.Drawing.Size(120, 23);
			this.BtnGetUser.TabIndex = 23;
			this.BtnGetUser.Text = "GetUser";
			this.BtnGetUser.Click += new System.EventHandler(this.BtnGetUser_Click);
			// 
			// LblUserId
			// 
			this.LblUserId.Location = new System.Drawing.Point(32, 16);
			this.LblUserId.Name = "LblUserId";
			this.LblUserId.Size = new System.Drawing.Size(64, 23);
			this.LblUserId.TabIndex = 24;
			this.LblUserId.Text = "User Id:";
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.TxtNew);
			this.GrpResult.Controls.Add(this.LblNew);
			this.GrpResult.Controls.Add(this.TxtChanged);
			this.GrpResult.Controls.Add(this.LblLastChanged);
			this.GrpResult.Controls.Add(this.TxtVerified);
			this.GrpResult.Controls.Add(this.LblVerified);
			this.GrpResult.Controls.Add(this.TxtUserIdRet);
			this.GrpResult.Controls.Add(this.LblUserIdRet);
			this.GrpResult.Controls.Add(this.TxtStar);
			this.GrpResult.Controls.Add(this.LblStar);
			this.GrpResult.Controls.Add(this.TxtStoreUrl);
			this.GrpResult.Controls.Add(this.LblStoreUrl);
			this.GrpResult.Controls.Add(this.TxtSite);
			this.GrpResult.Controls.Add(this.LblSite);
			this.GrpResult.Controls.Add(this.TxtRegDate);
			this.GrpResult.Controls.Add(this.TxtEmail);
			this.GrpResult.Controls.Add(this.TxtFeedBackScore);
			this.GrpResult.Controls.Add(this.TxtSellerLevel);
			this.GrpResult.Controls.Add(this.LblEmail);
			this.GrpResult.Controls.Add(this.LblRegDate);
			this.GrpResult.Controls.Add(this.LblFeedBackScore);
			this.GrpResult.Controls.Add(this.LblSellerLevel);
			this.GrpResult.Location = new System.Drawing.Point(8, 80);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(408, 208);
			this.GrpResult.TabIndex = 25;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// TxtNew
			// 
			this.TxtNew.Location = new System.Drawing.Point(328, 80);
			this.TxtNew.Name = "TxtNew";
			this.TxtNew.ReadOnly = true;
			this.TxtNew.Size = new System.Drawing.Size(72, 20);
			this.TxtNew.TabIndex = 60;
			this.TxtNew.Text = "";
			// 
			// LblNew
			// 
			this.LblNew.Location = new System.Drawing.Point(216, 80);
			this.LblNew.Name = "LblNew";
			this.LblNew.Size = new System.Drawing.Size(112, 23);
			this.LblNew.TabIndex = 59;
			this.LblNew.Text = "New User:";
			// 
			// TxtChanged
			// 
			this.TxtChanged.Location = new System.Drawing.Point(128, 176);
			this.TxtChanged.Name = "TxtChanged";
			this.TxtChanged.ReadOnly = true;
			this.TxtChanged.Size = new System.Drawing.Size(136, 20);
			this.TxtChanged.TabIndex = 57;
			this.TxtChanged.Text = "";
			// 
			// LblLastChanged
			// 
			this.LblLastChanged.Location = new System.Drawing.Point(16, 176);
			this.LblLastChanged.Name = "LblLastChanged";
			this.LblLastChanged.Size = new System.Drawing.Size(112, 23);
			this.LblLastChanged.TabIndex = 58;
			this.LblLastChanged.Text = "User Id Changed:";
			// 
			// TxtVerified
			// 
			this.TxtVerified.Location = new System.Drawing.Point(328, 104);
			this.TxtVerified.Name = "TxtVerified";
			this.TxtVerified.ReadOnly = true;
			this.TxtVerified.Size = new System.Drawing.Size(72, 20);
			this.TxtVerified.TabIndex = 56;
			this.TxtVerified.Text = "";
			// 
			// LblVerified
			// 
			this.LblVerified.Location = new System.Drawing.Point(216, 104);
			this.LblVerified.Name = "LblVerified";
			this.LblVerified.Size = new System.Drawing.Size(112, 23);
			this.LblVerified.TabIndex = 55;
			this.LblVerified.Text = "Verified:";
			// 
			// TxtUserIdRet
			// 
			this.TxtUserIdRet.Location = new System.Drawing.Point(128, 32);
			this.TxtUserIdRet.Name = "TxtUserIdRet";
			this.TxtUserIdRet.ReadOnly = true;
			this.TxtUserIdRet.Size = new System.Drawing.Size(72, 20);
			this.TxtUserIdRet.TabIndex = 54;
			this.TxtUserIdRet.Text = "";
			// 
			// LblUserIdRet
			// 
			this.LblUserIdRet.Location = new System.Drawing.Point(16, 32);
			this.LblUserIdRet.Name = "LblUserIdRet";
			this.LblUserIdRet.Size = new System.Drawing.Size(112, 23);
			this.LblUserIdRet.TabIndex = 53;
			this.LblUserIdRet.Text = "User Id:";
			// 
			// TxtStar
			// 
			this.TxtStar.Location = new System.Drawing.Point(128, 104);
			this.TxtStar.Name = "TxtStar";
			this.TxtStar.ReadOnly = true;
			this.TxtStar.Size = new System.Drawing.Size(72, 20);
			this.TxtStar.TabIndex = 51;
			this.TxtStar.Text = "";
			// 
			// LblStar
			// 
			this.LblStar.Location = new System.Drawing.Point(16, 104);
			this.LblStar.Name = "LblStar";
			this.LblStar.Size = new System.Drawing.Size(112, 23);
			this.LblStar.TabIndex = 52;
			this.LblStar.Text = "Star:";
			// 
			// TxtStoreUrl
			// 
			this.TxtStoreUrl.Location = new System.Drawing.Point(128, 128);
			this.TxtStoreUrl.Name = "TxtStoreUrl";
			this.TxtStoreUrl.ReadOnly = true;
			this.TxtStoreUrl.Size = new System.Drawing.Size(272, 20);
			this.TxtStoreUrl.TabIndex = 49;
			this.TxtStoreUrl.Text = "";
			// 
			// LblStoreUrl
			// 
			this.LblStoreUrl.Location = new System.Drawing.Point(16, 128);
			this.LblStoreUrl.Name = "LblStoreUrl";
			this.LblStoreUrl.Size = new System.Drawing.Size(112, 23);
			this.LblStoreUrl.TabIndex = 50;
			this.LblStoreUrl.Text = "Store Url:";
			// 
			// TxtSite
			// 
			this.TxtSite.Location = new System.Drawing.Point(328, 56);
			this.TxtSite.Name = "TxtSite";
			this.TxtSite.ReadOnly = true;
			this.TxtSite.Size = new System.Drawing.Size(72, 20);
			this.TxtSite.TabIndex = 47;
			this.TxtSite.Text = "";
			// 
			// LblSite
			// 
			this.LblSite.Location = new System.Drawing.Point(216, 56);
			this.LblSite.Name = "LblSite";
			this.LblSite.Size = new System.Drawing.Size(112, 23);
			this.LblSite.TabIndex = 48;
			this.LblSite.Text = "Site:";
			// 
			// TxtRegDate
			// 
			this.TxtRegDate.Location = new System.Drawing.Point(128, 152);
			this.TxtRegDate.Name = "TxtRegDate";
			this.TxtRegDate.ReadOnly = true;
			this.TxtRegDate.Size = new System.Drawing.Size(136, 20);
			this.TxtRegDate.TabIndex = 45;
			this.TxtRegDate.Text = "";
			// 
			// TxtEmail
			// 
			this.TxtEmail.Location = new System.Drawing.Point(128, 56);
			this.TxtEmail.Name = "TxtEmail";
			this.TxtEmail.ReadOnly = true;
			this.TxtEmail.Size = new System.Drawing.Size(72, 20);
			this.TxtEmail.TabIndex = 40;
			this.TxtEmail.Text = "";
			// 
			// TxtFeedBackScore
			// 
			this.TxtFeedBackScore.Location = new System.Drawing.Point(128, 80);
			this.TxtFeedBackScore.Name = "TxtFeedBackScore";
			this.TxtFeedBackScore.ReadOnly = true;
			this.TxtFeedBackScore.Size = new System.Drawing.Size(72, 20);
			this.TxtFeedBackScore.TabIndex = 39;
			this.TxtFeedBackScore.Text = "";
			// 
			// TxtSellerLevel
			// 
			this.TxtSellerLevel.Location = new System.Drawing.Point(328, 32);
			this.TxtSellerLevel.Name = "TxtSellerLevel";
			this.TxtSellerLevel.ReadOnly = true;
			this.TxtSellerLevel.Size = new System.Drawing.Size(72, 20);
			this.TxtSellerLevel.TabIndex = 41;
			this.TxtSellerLevel.Text = "";
			// 
			// LblEmail
			// 
			this.LblEmail.Location = new System.Drawing.Point(16, 56);
			this.LblEmail.Name = "LblEmail";
			this.LblEmail.Size = new System.Drawing.Size(112, 23);
			this.LblEmail.TabIndex = 43;
			this.LblEmail.Text = "Email:";
			// 
			// LblRegDate
			// 
			this.LblRegDate.Location = new System.Drawing.Point(16, 152);
			this.LblRegDate.Name = "LblRegDate";
			this.LblRegDate.Size = new System.Drawing.Size(112, 23);
			this.LblRegDate.TabIndex = 46;
			this.LblRegDate.Text = "Registration Date:";
			// 
			// LblFeedBackScore
			// 
			this.LblFeedBackScore.Location = new System.Drawing.Point(16, 80);
			this.LblFeedBackScore.Name = "LblFeedBackScore";
			this.LblFeedBackScore.Size = new System.Drawing.Size(112, 23);
			this.LblFeedBackScore.TabIndex = 44;
			this.LblFeedBackScore.Text = "Feedback Score:";
			// 
			// LblSellerLevel
			// 
			this.LblSellerLevel.Location = new System.Drawing.Point(216, 32);
			this.LblSellerLevel.Name = "LblSellerLevel";
			this.LblSellerLevel.Size = new System.Drawing.Size(112, 23);
			this.LblSellerLevel.TabIndex = 42;
			this.LblSellerLevel.Text = "Seller Level:";
			// 
			// TxtItemId
			// 
			this.TxtItemId.Location = new System.Drawing.Point(256, 16);
			this.TxtItemId.Name = "TxtItemId";
			this.TxtItemId.Size = new System.Drawing.Size(80, 20);
			this.TxtItemId.TabIndex = 26;
			this.TxtItemId.Text = "";
			// 
			// LblItemId
			// 
			this.LblItemId.Location = new System.Drawing.Point(192, 16);
			this.LblItemId.Name = "LblItemId";
			this.LblItemId.Size = new System.Drawing.Size(64, 23);
			this.LblItemId.TabIndex = 27;
			this.LblItemId.Text = "Item Id:";
			// 
			// FrmGetUser
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(424, 293);
			this.Controls.Add(this.TxtItemId);
			this.Controls.Add(this.LblItemId);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.TxtUserId);
			this.Controls.Add(this.BtnGetUser);
			this.Controls.Add(this.LblUserId);
			this.Name = "FrmGetUser";
			this.Text = "GetUser";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetUser_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtUserIdRet.Text = "";
				TxtEmail.Text = "";
				TxtFeedBackScore.Text = "";
				TxtRegDate.Text = "";
				TxtSellerLevel.Text = "";
				TxtSite.Text = "";
				TxtStar.Text = "";
				TxtStoreUrl.Text = "";
				TxtNew.Text = "";
				TxtVerified.Text = "";
				TxtChanged.Text = "";

				
				GetUserCall apicall = new GetUserCall(Context);
	
				if (TxtUserId.Text != String.Empty)
				apicall.UserID = TxtUserId.Text;

				if (TxtItemId.Text != String.Empty)
					apicall.ItemID = TxtItemId.Text;

				UserType user = apicall.GetUser();
	
				TxtUserIdRet.Text = user.UserID;
				TxtEmail.Text = user.Email;
				TxtFeedBackScore.Text = user.FeedbackScore.ToString();
				TxtRegDate.Text = user.RegistrationDate.ToString();
				TxtSellerLevel.Text = user.SellerInfo.SellerLevel.ToString();
				TxtSite.Text = user.Site.ToString();
				TxtStar.Text = user.FeedbackRatingStar.ToString();
				TxtStoreUrl.Text = user.SellerInfo.StoreURL;
				TxtNew.Text = user.NewUser.ToString();
				TxtVerified.Text = user.IDVerified.ToString();
				TxtChanged.Text = user.UserIDLastChanged.ToString();

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

	}
}
