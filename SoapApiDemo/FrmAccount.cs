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
using eBay.Service.Call;
using eBay.Service.Core.Soap;
using eBay.Service.Core.Sdk;
using System.Xml;
using eBay.Service.Util;
using Samples.Helper;


namespace SoapLibraryDemo
{

	/// <summary>
	/// Summary description for AccountForm.
	/// </summary>
	public class FrmAccount : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button BtnContinue;
		private System.Windows.Forms.Button BtnCancel;
		private System.Windows.Forms.GroupBox GrpUrl;
		private System.Windows.Forms.Label LblUrl;
		private System.Windows.Forms.TextBox TxtUrl;
		private System.Windows.Forms.GroupBox GrpAccountAuth;
		private System.Windows.Forms.GroupBox GrpProgramAuth;
		private System.Windows.Forms.Label LblDeveloper;
		private System.Windows.Forms.Label LblApplication;
		private System.Windows.Forms.TextBox TxtApplication;
		private System.Windows.Forms.TextBox TxtDeveloper;
		private System.Windows.Forms.TextBox TxtCertificate;
		private System.Windows.Forms.Label LblCertificate;
		private System.Windows.Forms.TabControl TabSettings;
		private System.Windows.Forms.TabPage TabPageCredentials;
		private System.Windows.Forms.TabPage TabPageToken;
		private System.Windows.Forms.Label LblToken;
		internal System.Windows.Forms.Button BtnGenerateSID;
		internal System.Windows.Forms.Label LblThree;
		internal System.Windows.Forms.Label LblOne;
		internal System.Windows.Forms.Button BtnLaunchBrowserWithSecret;
		internal System.Windows.Forms.TextBox TxtSessionID;
		internal System.Windows.Forms.Label LblSessionID;
		internal System.Windows.Forms.Label LblRuInstruct;
		internal System.Windows.Forms.Label LblSidLength;
		internal System.Windows.Forms.Label LblTwo;
		private System.Windows.Forms.GroupBox GrpFetchToken;
		internal System.Windows.Forms.Label LblFetch;
		private System.Windows.Forms.Button BtnFetchToken;
		private System.Windows.Forms.TextBox TxtToken;
		private System.Windows.Forms.Label LblSignInUrl;
		private System.Windows.Forms.TextBox TxtSignInUrl;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.ColumnHeader columnHeader2;
		private System.Windows.Forms.ColumnHeader columnHeader3;
		private System.Windows.Forms.Label LblEPSUrl;
		private System.Windows.Forms.TextBox TxtEPSUrl;
		private System.Windows.Forms.GroupBox GrpLogging;
		private System.Windows.Forms.CheckBox ChkEnableLog;
		private System.Windows.Forms.CheckBox ChkLogMessages;
		private System.Windows.Forms.CheckBox ChkLogExceptions;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox TxtLogFile;
		internal System.Windows.Forms.Button BtnLogFile;
		private System.Windows.Forms.SaveFileDialog DlgSaveLog;
		private System.Windows.Forms.TabPage TabOther;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label LblSite;
		private System.Windows.Forms.Label LblVersion;
		private System.Windows.Forms.Label LblTimeOut;
		private System.Windows.Forms.Label LblLanguage;
		private System.Windows.Forms.TextBox TxtVersion;
		private System.Windows.Forms.TextBox TxtTimeOut;
		private System.Windows.Forms.ComboBox CboLanguage;
		private System.Windows.Forms.ComboBox CboSite;
		private System.Windows.Forms.TextBox txtRulName;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmAccount()
		{
			InitializeComponent();
			SetFormStartPosition();
		}

		private void SetFormStartPosition()
		{
			this.StartPosition = FormStartPosition.CenterParent;
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
			this.BtnContinue = new System.Windows.Forms.Button();
			this.BtnCancel = new System.Windows.Forms.Button();
			this.TabSettings = new System.Windows.Forms.TabControl();
			this.TabPageCredentials = new System.Windows.Forms.TabPage();
			this.GrpUrl = new System.Windows.Forms.GroupBox();
			this.LblEPSUrl = new System.Windows.Forms.Label();
			this.TxtEPSUrl = new System.Windows.Forms.TextBox();
			this.LblSignInUrl = new System.Windows.Forms.Label();
			this.TxtSignInUrl = new System.Windows.Forms.TextBox();
			this.LblUrl = new System.Windows.Forms.Label();
			this.TxtUrl = new System.Windows.Forms.TextBox();
			this.GrpAccountAuth = new System.Windows.Forms.GroupBox();
			this.TxtToken = new System.Windows.Forms.TextBox();
			this.LblToken = new System.Windows.Forms.Label();
			this.GrpProgramAuth = new System.Windows.Forms.GroupBox();
			this.LblDeveloper = new System.Windows.Forms.Label();
			this.LblApplication = new System.Windows.Forms.Label();
			this.TxtApplication = new System.Windows.Forms.TextBox();
			this.TxtDeveloper = new System.Windows.Forms.TextBox();
			this.TxtCertificate = new System.Windows.Forms.TextBox();
			this.LblCertificate = new System.Windows.Forms.Label();
			this.TabOther = new System.Windows.Forms.TabPage();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.CboSite = new System.Windows.Forms.ComboBox();
			this.CboLanguage = new System.Windows.Forms.ComboBox();
			this.TxtTimeOut = new System.Windows.Forms.TextBox();
			this.LblTimeOut = new System.Windows.Forms.Label();
			this.LblSite = new System.Windows.Forms.Label();
			this.LblLanguage = new System.Windows.Forms.Label();
			this.TxtVersion = new System.Windows.Forms.TextBox();
			this.LblVersion = new System.Windows.Forms.Label();
			this.GrpLogging = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.BtnLogFile = new System.Windows.Forms.Button();
			this.TxtLogFile = new System.Windows.Forms.TextBox();
			this.ChkLogExceptions = new System.Windows.Forms.CheckBox();
			this.ChkLogMessages = new System.Windows.Forms.CheckBox();
			this.ChkEnableLog = new System.Windows.Forms.CheckBox();
			this.TabPageToken = new System.Windows.Forms.TabPage();
			this.GrpFetchToken = new System.Windows.Forms.GroupBox();
			this.txtRulName = new System.Windows.Forms.TextBox();
			this.BtnFetchToken = new System.Windows.Forms.Button();
			this.LblTwo = new System.Windows.Forms.Label();
			this.BtnGenerateSID = new System.Windows.Forms.Button();
			this.LblThree = new System.Windows.Forms.Label();
			this.LblFetch = new System.Windows.Forms.Label();
			this.LblOne = new System.Windows.Forms.Label();
			this.BtnLaunchBrowserWithSecret = new System.Windows.Forms.Button();
			this.TxtSessionID = new System.Windows.Forms.TextBox();
			this.LblSessionID = new System.Windows.Forms.Label();
			this.LblRuInstruct = new System.Windows.Forms.Label();
			this.LblSidLength = new System.Windows.Forms.Label();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
			this.DlgSaveLog = new System.Windows.Forms.SaveFileDialog();
			this.TabSettings.SuspendLayout();
			this.TabPageCredentials.SuspendLayout();
			this.GrpUrl.SuspendLayout();
			this.GrpAccountAuth.SuspendLayout();
			this.GrpProgramAuth.SuspendLayout();
			this.TabOther.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.GrpLogging.SuspendLayout();
			this.TabPageToken.SuspendLayout();
			this.GrpFetchToken.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnContinue
			// 
			this.BtnContinue.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.BtnContinue.Location = new System.Drawing.Point(176, 448);
			this.BtnContinue.Name = "BtnContinue";
			this.BtnContinue.Size = new System.Drawing.Size(88, 25);
			this.BtnContinue.TabIndex = 38;
			this.BtnContinue.Text = "Continue";
			this.BtnContinue.Click += new System.EventHandler(this.BtnContinue_Click);
			// 
			// BtnCancel
			// 
			this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.BtnCancel.Location = new System.Drawing.Point(280, 448);
			this.BtnCancel.Name = "BtnCancel";
			this.BtnCancel.Size = new System.Drawing.Size(88, 25);
			this.BtnCancel.TabIndex = 39;
			this.BtnCancel.Text = "Cancel";
			// 
			// TabSettings
			// 
			this.TabSettings.Controls.Add(this.TabPageCredentials);
			this.TabSettings.Controls.Add(this.TabOther);
			this.TabSettings.Controls.Add(this.TabPageToken);
			this.TabSettings.Location = new System.Drawing.Point(8, 8);
			this.TabSettings.Name = "TabSettings";
			this.TabSettings.SelectedIndex = 0;
			this.TabSettings.Size = new System.Drawing.Size(360, 432);
			this.TabSettings.TabIndex = 40;
			// 
			// TabPageCredentials
			// 
			this.TabPageCredentials.Controls.Add(this.GrpUrl);
			this.TabPageCredentials.Controls.Add(this.GrpAccountAuth);
			this.TabPageCredentials.Controls.Add(this.GrpProgramAuth);
			this.TabPageCredentials.Location = new System.Drawing.Point(4, 22);
			this.TabPageCredentials.Name = "TabPageCredentials";
			this.TabPageCredentials.Size = new System.Drawing.Size(352, 406);
			this.TabPageCredentials.TabIndex = 0;
			this.TabPageCredentials.Text = "Credentials";
			// 
			// GrpUrl
			// 
			this.GrpUrl.Controls.Add(this.LblEPSUrl);
			this.GrpUrl.Controls.Add(this.TxtEPSUrl);
			this.GrpUrl.Controls.Add(this.LblSignInUrl);
			this.GrpUrl.Controls.Add(this.TxtSignInUrl);
			this.GrpUrl.Controls.Add(this.LblUrl);
			this.GrpUrl.Controls.Add(this.TxtUrl);
			this.GrpUrl.Location = new System.Drawing.Point(16, 296);
			this.GrpUrl.Name = "GrpUrl";
			this.GrpUrl.Size = new System.Drawing.Size(328, 104);
			this.GrpUrl.TabIndex = 44;
			this.GrpUrl.TabStop = false;
			this.GrpUrl.Text = "Urls to make eBay API Call to:";
			// 
			// LblEPSUrl
			// 
			this.LblEPSUrl.Location = new System.Drawing.Point(24, 72);
			this.LblEPSUrl.Name = "LblEPSUrl";
			this.LblEPSUrl.Size = new System.Drawing.Size(104, 19);
			this.LblEPSUrl.TabIndex = 29;
			this.LblEPSUrl.Text = "EPS Url:";
			// 
			// TxtEPSUrl
			// 
			this.TxtEPSUrl.Location = new System.Drawing.Point(128, 72);
			this.TxtEPSUrl.Name = "TxtEPSUrl";
			this.TxtEPSUrl.Size = new System.Drawing.Size(178, 20);
			this.TxtEPSUrl.TabIndex = 28;
			this.TxtEPSUrl.Text = "";
			// 
			// LblSignInUrl
			// 
			this.LblSignInUrl.Location = new System.Drawing.Point(24, 48);
			this.LblSignInUrl.Name = "LblSignInUrl";
			this.LblSignInUrl.Size = new System.Drawing.Size(104, 19);
			this.LblSignInUrl.TabIndex = 27;
			this.LblSignInUrl.Text = "Sign-In Url:";
			// 
			// TxtSignInUrl
			// 
			this.TxtSignInUrl.Location = new System.Drawing.Point(128, 48);
			this.TxtSignInUrl.Name = "TxtSignInUrl";
			this.TxtSignInUrl.Size = new System.Drawing.Size(178, 20);
			this.TxtSignInUrl.TabIndex = 26;
			this.TxtSignInUrl.Text = "";
			// 
			// LblUrl
			// 
			this.LblUrl.Location = new System.Drawing.Point(24, 24);
			this.LblUrl.Name = "LblUrl";
			this.LblUrl.Size = new System.Drawing.Size(104, 19);
			this.LblUrl.TabIndex = 25;
			this.LblUrl.Text = "API Server Url:";
			// 
			// TxtUrl
			// 
			this.TxtUrl.Location = new System.Drawing.Point(128, 24);
			this.TxtUrl.Name = "TxtUrl";
			this.TxtUrl.Size = new System.Drawing.Size(178, 20);
			this.TxtUrl.TabIndex = 0;
			this.TxtUrl.Text = "";
			// 
			// GrpAccountAuth
			// 
			this.GrpAccountAuth.Controls.Add(this.TxtToken);
			this.GrpAccountAuth.Controls.Add(this.LblToken);
			this.GrpAccountAuth.Location = new System.Drawing.Point(16, 128);
			this.GrpAccountAuth.Name = "GrpAccountAuth";
			this.GrpAccountAuth.Size = new System.Drawing.Size(328, 160);
			this.GrpAccountAuth.TabIndex = 43;
			this.GrpAccountAuth.TabStop = false;
			this.GrpAccountAuth.Text = "eBay Seller Account:";
			// 
			// TxtToken
			// 
			this.TxtToken.Location = new System.Drawing.Point(32, 32);
			this.TxtToken.Multiline = true;
			this.TxtToken.Name = "TxtToken";
			this.TxtToken.Size = new System.Drawing.Size(272, 120);
			this.TxtToken.TabIndex = 45;
			this.TxtToken.Text = "";
			// 
			// LblToken
			// 
			this.LblToken.Location = new System.Drawing.Point(16, 16);
			this.LblToken.Name = "LblToken";
			this.LblToken.Size = new System.Drawing.Size(100, 16);
			this.LblToken.TabIndex = 44;
			this.LblToken.Text = "Token:";
			// 
			// GrpProgramAuth
			// 
			this.GrpProgramAuth.Controls.Add(this.LblDeveloper);
			this.GrpProgramAuth.Controls.Add(this.LblApplication);
			this.GrpProgramAuth.Controls.Add(this.TxtApplication);
			this.GrpProgramAuth.Controls.Add(this.TxtDeveloper);
			this.GrpProgramAuth.Controls.Add(this.TxtCertificate);
			this.GrpProgramAuth.Controls.Add(this.LblCertificate);
			this.GrpProgramAuth.Location = new System.Drawing.Point(16, 16);
			this.GrpProgramAuth.Name = "GrpProgramAuth";
			this.GrpProgramAuth.Size = new System.Drawing.Size(328, 104);
			this.GrpProgramAuth.TabIndex = 41;
			this.GrpProgramAuth.TabStop = false;
			this.GrpProgramAuth.Text = "API Certificate";
			// 
			// LblDeveloper
			// 
			this.LblDeveloper.Location = new System.Drawing.Point(16, 25);
			this.LblDeveloper.Name = "LblDeveloper";
			this.LblDeveloper.Size = new System.Drawing.Size(112, 19);
			this.LblDeveloper.TabIndex = 18;
			this.LblDeveloper.Text = "Developer:";
			// 
			// LblApplication
			// 
			this.LblApplication.Location = new System.Drawing.Point(16, 49);
			this.LblApplication.Name = "LblApplication";
			this.LblApplication.Size = new System.Drawing.Size(112, 19);
			this.LblApplication.TabIndex = 20;
			this.LblApplication.Text = "Application:";
			// 
			// TxtApplication
			// 
			this.TxtApplication.Location = new System.Drawing.Point(128, 48);
			this.TxtApplication.Name = "TxtApplication";
			this.TxtApplication.Size = new System.Drawing.Size(178, 20);
			this.TxtApplication.TabIndex = 1;
			this.TxtApplication.Text = "";
			// 
			// TxtDeveloper
			// 
			this.TxtDeveloper.Location = new System.Drawing.Point(128, 24);
			this.TxtDeveloper.Name = "TxtDeveloper";
			this.TxtDeveloper.Size = new System.Drawing.Size(178, 20);
			this.TxtDeveloper.TabIndex = 0;
			this.TxtDeveloper.Text = "";
			// 
			// TxtCertificate
			// 
			this.TxtCertificate.Location = new System.Drawing.Point(128, 72);
			this.TxtCertificate.Name = "TxtCertificate";
			this.TxtCertificate.PasswordChar = '*';
			this.TxtCertificate.Size = new System.Drawing.Size(178, 20);
			this.TxtCertificate.TabIndex = 2;
			this.TxtCertificate.Text = "";
			// 
			// LblCertificate
			// 
			this.LblCertificate.Location = new System.Drawing.Point(16, 73);
			this.LblCertificate.Name = "LblCertificate";
			this.LblCertificate.Size = new System.Drawing.Size(112, 19);
			this.LblCertificate.TabIndex = 22;
			this.LblCertificate.Text = "Certificate:";
			// 
			// TabOther
			// 
			this.TabOther.Controls.Add(this.groupBox1);
			this.TabOther.Controls.Add(this.GrpLogging);
			this.TabOther.Location = new System.Drawing.Point(4, 22);
			this.TabOther.Name = "TabOther";
			this.TabOther.Size = new System.Drawing.Size(352, 406);
			this.TabOther.TabIndex = 3;
			this.TabOther.Text = "Other";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.CboSite);
			this.groupBox1.Controls.Add(this.CboLanguage);
			this.groupBox1.Controls.Add(this.TxtTimeOut);
			this.groupBox1.Controls.Add(this.LblTimeOut);
			this.groupBox1.Controls.Add(this.LblSite);
			this.groupBox1.Controls.Add(this.LblLanguage);
			this.groupBox1.Controls.Add(this.TxtVersion);
			this.groupBox1.Controls.Add(this.LblVersion);
			this.groupBox1.Location = new System.Drawing.Point(8, 208);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(336, 128);
			this.groupBox1.TabIndex = 42;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "API Certificate";
			// 
			// CboSite
			// 
			this.CboSite.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboSite.Location = new System.Drawing.Point(128, 24);
			this.CboSite.Name = "CboSite";
			this.CboSite.Size = new System.Drawing.Size(176, 21);
			this.CboSite.TabIndex = 57;
			this.CboSite.SelectedIndexChanged += new System.EventHandler(this.CboSite_SelectedIndexChanged);
			// 
			// CboLanguage
			// 
			this.CboLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.CboLanguage.Location = new System.Drawing.Point(128, 48);
			this.CboLanguage.Name = "CboLanguage";
			this.CboLanguage.Size = new System.Drawing.Size(176, 21);
			this.CboLanguage.TabIndex = 56;
			this.CboLanguage.SelectedIndexChanged += new System.EventHandler(this.CboLanguage_SelectedIndexChanged);
			// 
			// TxtTimeOut
			// 
			this.TxtTimeOut.Location = new System.Drawing.Point(128, 96);
			this.TxtTimeOut.Name = "TxtTimeOut";
			this.TxtTimeOut.Size = new System.Drawing.Size(178, 20);
			this.TxtTimeOut.TabIndex = 23;
			this.TxtTimeOut.Text = "";
			// 
			// LblTimeOut
			// 
			this.LblTimeOut.Location = new System.Drawing.Point(16, 96);
			this.LblTimeOut.Name = "LblTimeOut";
			this.LblTimeOut.Size = new System.Drawing.Size(112, 16);
			this.LblTimeOut.TabIndex = 24;
			this.LblTimeOut.Text = "Timeout (ms):";
			// 
			// LblSite
			// 
			this.LblSite.Location = new System.Drawing.Point(16, 24);
			this.LblSite.Name = "LblSite";
			this.LblSite.Size = new System.Drawing.Size(112, 16);
			this.LblSite.TabIndex = 18;
			this.LblSite.Text = "Site:";
			// 
			// LblLanguage
			// 
			this.LblLanguage.Location = new System.Drawing.Point(16, 48);
			this.LblLanguage.Name = "LblLanguage";
			this.LblLanguage.Size = new System.Drawing.Size(112, 16);
			this.LblLanguage.TabIndex = 20;
			this.LblLanguage.Text = "Error Language:";
			// 
			// TxtVersion
			// 
			this.TxtVersion.Location = new System.Drawing.Point(128, 72);
			this.TxtVersion.Name = "TxtVersion";
			this.TxtVersion.Size = new System.Drawing.Size(178, 20);
			this.TxtVersion.TabIndex = 2;
			this.TxtVersion.Text = "";
			// 
			// LblVersion
			// 
			this.LblVersion.Location = new System.Drawing.Point(16, 72);
			this.LblVersion.Name = "LblVersion";
			this.LblVersion.Size = new System.Drawing.Size(112, 16);
			this.LblVersion.TabIndex = 22;
			this.LblVersion.Text = "Version:";
			// 
			// GrpLogging
			// 
			this.GrpLogging.Controls.Add(this.label1);
			this.GrpLogging.Controls.Add(this.BtnLogFile);
			this.GrpLogging.Controls.Add(this.TxtLogFile);
			this.GrpLogging.Controls.Add(this.ChkLogExceptions);
			this.GrpLogging.Controls.Add(this.ChkLogMessages);
			this.GrpLogging.Controls.Add(this.ChkEnableLog);
			this.GrpLogging.Location = new System.Drawing.Point(8, 8);
			this.GrpLogging.Name = "GrpLogging";
			this.GrpLogging.Size = new System.Drawing.Size(336, 192);
			this.GrpLogging.TabIndex = 0;
			this.GrpLogging.TabStop = false;
			this.GrpLogging.Text = "Configure Logging";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 96);
			this.label1.Name = "label1";
			this.label1.TabIndex = 64;
			this.label1.Text = "Log File";
			// 
			// BtnLogFile
			// 
			this.BtnLogFile.Location = new System.Drawing.Point(216, 152);
			this.BtnLogFile.Name = "BtnLogFile";
			this.BtnLogFile.Size = new System.Drawing.Size(112, 23);
			this.BtnLogFile.TabIndex = 63;
			this.BtnLogFile.Text = "Browse";
			this.BtnLogFile.Click += new System.EventHandler(this.BtnLogFile_Click);
			// 
			// TxtLogFile
			// 
			this.TxtLogFile.Location = new System.Drawing.Point(16, 120);
			this.TxtLogFile.Name = "TxtLogFile";
			this.TxtLogFile.Size = new System.Drawing.Size(312, 20);
			this.TxtLogFile.TabIndex = 3;
			this.TxtLogFile.Text = "Log.txt";
			// 
			// ChkLogExceptions
			// 
			this.ChkLogExceptions.Location = new System.Drawing.Point(40, 72);
			this.ChkLogExceptions.Name = "ChkLogExceptions";
			this.ChkLogExceptions.TabIndex = 2;
			this.ChkLogExceptions.Text = "Log Exceptions";
			// 
			// ChkLogMessages
			// 
			this.ChkLogMessages.Location = new System.Drawing.Point(40, 48);
			this.ChkLogMessages.Name = "ChkLogMessages";
			this.ChkLogMessages.TabIndex = 1;
			this.ChkLogMessages.Text = "Log Messages";
			// 
			// ChkEnableLog
			// 
			this.ChkEnableLog.Checked = true;
			this.ChkEnableLog.CheckState = System.Windows.Forms.CheckState.Checked;
			this.ChkEnableLog.Location = new System.Drawing.Point(16, 24);
			this.ChkEnableLog.Name = "ChkEnableLog";
			this.ChkEnableLog.TabIndex = 0;
			this.ChkEnableLog.Text = "Enable Logging";
			this.ChkEnableLog.CheckedChanged += new System.EventHandler(this.ChkEnableLog_CheckedChanged);
			// 
			// TabPageToken
			// 
			this.TabPageToken.Controls.Add(this.GrpFetchToken);
			this.TabPageToken.Location = new System.Drawing.Point(4, 22);
			this.TabPageToken.Name = "TabPageToken";
			this.TabPageToken.Size = new System.Drawing.Size(352, 406);
			this.TabPageToken.TabIndex = 1;
			this.TabPageToken.Text = "Token";
			// 
			// GrpFetchToken
			// 
			this.GrpFetchToken.Controls.Add(this.txtRulName);
			this.GrpFetchToken.Controls.Add(this.BtnFetchToken);
			this.GrpFetchToken.Controls.Add(this.LblTwo);
			this.GrpFetchToken.Controls.Add(this.BtnGenerateSID);
			this.GrpFetchToken.Controls.Add(this.LblThree);
			this.GrpFetchToken.Controls.Add(this.LblFetch);
			this.GrpFetchToken.Controls.Add(this.LblOne);
			this.GrpFetchToken.Controls.Add(this.BtnLaunchBrowserWithSecret);
			this.GrpFetchToken.Controls.Add(this.TxtSessionID);
			this.GrpFetchToken.Controls.Add(this.LblSessionID);
			this.GrpFetchToken.Controls.Add(this.LblRuInstruct);
			this.GrpFetchToken.Controls.Add(this.LblSidLength);
			this.GrpFetchToken.Location = new System.Drawing.Point(8, 8);
			this.GrpFetchToken.Name = "GrpFetchToken";
			this.GrpFetchToken.Size = new System.Drawing.Size(336, 232);
			this.GrpFetchToken.TabIndex = 58;
			this.GrpFetchToken.TabStop = false;
			this.GrpFetchToken.Text = "Fetch Token";
			// 
			// txtRulName
			// 
			this.txtRulName.Location = new System.Drawing.Point(40, 48);
			this.txtRulName.Name = "txtRulName";
			this.txtRulName.Size = new System.Drawing.Size(148, 20);
			this.txtRulName.TabIndex = 73;
			this.txtRulName.Text = "";
			// 
			// BtnFetchToken
			// 
			this.BtnFetchToken.Location = new System.Drawing.Point(112, 200);
			this.BtnFetchToken.Name = "BtnFetchToken";
			this.BtnFetchToken.Size = new System.Drawing.Size(88, 24);
			this.BtnFetchToken.TabIndex = 70;
			this.BtnFetchToken.Text = "Fetch Token";
			this.BtnFetchToken.Click += new System.EventHandler(this.BtnFetchToken_Click);
			// 
			// LblTwo
			// 
			this.LblTwo.Location = new System.Drawing.Point(16, 72);
			this.LblTwo.Name = "LblTwo";
			this.LblTwo.Size = new System.Drawing.Size(21, 16);
			this.LblTwo.TabIndex = 68;
			this.LblTwo.Text = "2)";
			this.LblTwo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// BtnGenerateSID
			// 
			this.BtnGenerateSID.Location = new System.Drawing.Point(248, 120);
			this.BtnGenerateSID.Name = "BtnGenerateSID";
			this.BtnGenerateSID.Size = new System.Drawing.Size(64, 23);
			this.BtnGenerateSID.TabIndex = 61;
			this.BtnGenerateSID.Text = "Retrieve";
			this.BtnGenerateSID.Click += new System.EventHandler(this.BtnGenerateSID_Click);
			// 
			// LblThree
			// 
			this.LblThree.Location = new System.Drawing.Point(16, 208);
			this.LblThree.Name = "LblThree";
			this.LblThree.Size = new System.Drawing.Size(21, 16);
			this.LblThree.TabIndex = 66;
			this.LblThree.Text = "3)";
			this.LblThree.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// LblFetch
			// 
			this.LblFetch.Location = new System.Drawing.Point(40, 208);
			this.LblFetch.Name = "LblFetch";
			this.LblFetch.Size = new System.Drawing.Size(40, 16);
			this.LblFetch.TabIndex = 67;
			this.LblFetch.Text = "Click";
			// 
			// LblOne
			// 
			this.LblOne.Location = new System.Drawing.Point(16, 24);
			this.LblOne.Name = "LblOne";
			this.LblOne.Size = new System.Drawing.Size(21, 16);
			this.LblOne.TabIndex = 57;
			this.LblOne.Text = "1)";
			this.LblOne.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// BtnLaunchBrowserWithSecret
			// 
			this.BtnLaunchBrowserWithSecret.Location = new System.Drawing.Point(112, 160);
			this.BtnLaunchBrowserWithSecret.Name = "BtnLaunchBrowserWithSecret";
			this.BtnLaunchBrowserWithSecret.Size = new System.Drawing.Size(88, 23);
			this.BtnLaunchBrowserWithSecret.TabIndex = 65;
			this.BtnLaunchBrowserWithSecret.Text = "Launch Sign-In";
			this.BtnLaunchBrowserWithSecret.Click += new System.EventHandler(this.BtnLaunchBrowserWithSecret_Click);
			// 
			// TxtSessionID
			// 
			this.TxtSessionID.Location = new System.Drawing.Point(40, 120);
			this.TxtSessionID.Name = "TxtSessionID";
			this.TxtSessionID.Size = new System.Drawing.Size(148, 20);
			this.TxtSessionID.TabIndex = 59;
			this.TxtSessionID.Text = "";
			this.TxtSessionID.TextChanged += new System.EventHandler(this.TxtSessionID_TextChanged);
			// 
			// LblSessionID
			// 
			this.LblSessionID.Location = new System.Drawing.Point(40, 72);
			this.LblSessionID.Name = "LblSessionID";
			this.LblSessionID.Size = new System.Drawing.Size(260, 40);
			this.LblSessionID.TabIndex = 58;
			this.LblSessionID.Text = "Retrieve the Session ID to be associated with the token to be generated.  And lau" +
				"nch the Sign-In Page.";
			// 
			// LblRuInstruct
			// 
			this.LblRuInstruct.Location = new System.Drawing.Point(40, 24);
			this.LblRuInstruct.Name = "LblRuInstruct";
			this.LblRuInstruct.Size = new System.Drawing.Size(280, 16);
			this.LblRuInstruct.TabIndex = 62;
			this.LblRuInstruct.Text = "Please input an RuName.";
			// 
			// LblSidLength
			// 
			this.LblSidLength.Location = new System.Drawing.Point(200, 128);
			this.LblSidLength.Name = "LblSidLength";
			this.LblSidLength.Size = new System.Drawing.Size(36, 16);
			this.LblSidLength.TabIndex = 60;
			this.LblSidLength.Text = "0";
			this.LblSidLength.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "RuName";
			this.columnHeader1.Width = 132;
			// 
			// columnHeader2
			// 
			this.columnHeader2.Text = "AcceptURL";
			this.columnHeader2.Width = 75;
			// 
			// columnHeader3
			// 
			this.columnHeader3.Text = "RejectURL";
			this.columnHeader3.Width = 84;
			// 
			// DlgSaveLog
			// 
			this.DlgSaveLog.DefaultExt = "txt";
			this.DlgSaveLog.OverwritePrompt = false;
			this.DlgSaveLog.FileOk += new System.ComponentModel.CancelEventHandler(this.DlgSaveLog_FileOk);
			// 
			// FrmAccount
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(376, 478);
			this.Controls.Add(this.TabSettings);
			this.Controls.Add(this.BtnContinue);
			this.Controls.Add(this.BtnCancel);
			this.Name = "FrmAccount";
			this.Text = "Account Settings";
			this.Load += new System.EventHandler(this.FrmAccount_Load);
			this.TabSettings.ResumeLayout(false);
			this.TabPageCredentials.ResumeLayout(false);
			this.GrpUrl.ResumeLayout(false);
			this.GrpAccountAuth.ResumeLayout(false);
			this.GrpProgramAuth.ResumeLayout(false);
			this.TabOther.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.GrpLogging.ResumeLayout(false);
			this.TabPageToken.ResumeLayout(false);
			this.GrpFetchToken.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnContinue_Click(object sender, System.EventArgs e)
		{
			replaceConfigSettings();
		}

		private void FrmAccount_Load(object sender, System.EventArgs e)
		{
			if (Context == null)
			{
				Context = AppSettingHelper.GetApiContext();
				Context.ApiLogManager = new ApiLogManager();
				Context.ApiLogManager.ApiLoggerList.Add(new eBay.Service.Util.FileLogger("Log.txt", true, true, true));
				Context.ApiLogManager.EnableLogging = true;
				Context.Site = eBay.Service.Core.Soap.SiteCodeType.US;
			}
			string[] sites = Enum.GetNames(typeof(SiteCodeType));
			foreach (string site in sites)
			{
				if (site != "CustomCode")
				{
					CboSite.Items.Add(site);
				}
			}
			string[] langs = Enum.GetNames(typeof(ErrorLanguageCodeType));
			foreach (string lang in langs)
			{
				if (lang != "CustomCode")
				{
					CboLanguage.Items.Add(lang);
				}
			}
			CboSite.SelectedIndex = 0;

			SetBindings();
		}

		private void BtnGenerateSID_Click(object sender, System.EventArgs e)
		{
			try
			{
                if (!validateApiAccount(Context))
                {
                    MessageBox.Show("Pleasse fill in Api Account first.");
                    return;
                }

				GetSessionIDCall gsc = new GetSessionIDCall(Context);
				if(txtRulName.Text==string.Empty)
				{
					MessageBox.Show("Please input a RuName");
					return;
				}
				//the rule name can be retrived from the app.config file, otherwise the user must input it himself.
				
				Context.RuName = txtRulName.Text;
				String sessionID = gsc.GetSessionID(Context.RuName);
				TxtSessionID.Text = sessionID;
			} 
			catch(ApiException ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

        private bool validateApiAccount(ApiContext apiContext)
        {
            ApiAccount acc = apiContext.ApiCredential.ApiAccount;
            if (acc == null || acc.Application.Length == 0 || acc.Certificate.Length == 0 || acc.Developer.Length == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

		private void TxtSessionID_TextChanged(object sender, System.EventArgs e)
		{
			LblSidLength.Text = TxtSessionID.Text.Length.ToString();
		}

		private void LinkCredentials_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			this.TabSettings.SelectedTab = this.TabPageCredentials;
		}


		private void BtnFetchToken_Click(object sender, System.EventArgs e)
		{
			try
			{
				if(TxtSessionID.Text.Length == 0)
				{
					MessageBox.Show("Please retrieve a SessionID first, then launch Sign-In, before fetch token!");
					return;
				}
				FetchTokenCall ftc = new FetchTokenCall(Context);
				Context.ApiCredential.eBayToken = ftc.FetchToken(TxtSessionID.Text);
				this.TabSettings.SelectedTab = this.TabPageCredentials;
				TxtToken.Text = Context.ApiCredential.eBayToken;
				TxtToken.Focus();
			}
			catch (ApiException ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void BtnLaunchBrowserWithSecret_Click(object sender, System.EventArgs e)
		{
			try
			{
				if(this.TxtSessionID.Text.Length == 0)
				{
					MessageBox.Show("Please retrieve a SessionID first!");
					return;
				}
				SdkUtility.LaunchSignInPage(Context, TxtSessionID.Text);
			} 
			catch (SdkException ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void SetBindings()
		{
			TxtDeveloper.DataBindings.Clear();
			TxtApplication.DataBindings.Clear();
			TxtCertificate.DataBindings.Clear();
			TxtUrl.DataBindings.Clear();
			TxtSignInUrl.DataBindings.Clear();
			TxtEPSUrl.DataBindings.Clear();
			TxtVersion.DataBindings.Clear();
			TxtTimeOut.DataBindings.Clear();

			ChkEnableLog.DataBindings.Clear();
			ChkLogExceptions.DataBindings.Clear();
			ChkLogMessages.DataBindings.Clear();
			TxtLogFile.DataBindings.Clear();

			TxtDeveloper.DataBindings.Add("Text", Context.ApiCredential.ApiAccount, "Developer");
			TxtApplication.DataBindings.Add("Text", Context.ApiCredential.ApiAccount, "Application");
			TxtCertificate.DataBindings.Add("Text", Context.ApiCredential.ApiAccount, "Certificate");
			TxtToken.DataBindings.Add("Text", Context.ApiCredential, "eBayToken");
			TxtVersion.DataBindings.Add("Text", Context, "Version");
			TxtTimeOut.DataBindings.Add("Text", Context, "Timeout");
			txtRulName.DataBindings.Add("Text", Context, "RuleName");

			TxtUrl.DataBindings.Add("Text", Context, "SoapApiServerUrl");
			TxtSignInUrl.DataBindings.Add("Text", Context, "SignInUrl");
			TxtEPSUrl.DataBindings.Add("Text", Context, "EPSServerUrl");

			ChkEnableLog.DataBindings.Add("Checked", Context.ApiLogManager, "EnableLogging");
			ChkLogMessages.DataBindings.Add("Checked", Context.ApiLogManager, "LogApiMessages");
			ChkLogExceptions.DataBindings.Add("Checked", Context.ApiLogManager, "LogExceptions");

			eBay.Service.Util.FileLogger logfile = (eBay.Service.Util.FileLogger) Context.ApiLogManager.ApiLoggerList.ItemAt(0);
			TxtLogFile.DataBindings.Add("Text", logfile, "FileName");

			
		}


		private void ChkEnableLog_CheckedChanged(object sender, System.EventArgs e)
		{
			if (((CheckBox)sender).Checked)
			{
				ChkLogMessages.Enabled = true;
				ChkLogExceptions.Enabled = true;
				TxtLogFile.Enabled = true;
				BtnLogFile.Enabled = true;

			} 
			else
			{
				ChkLogMessages.Enabled = false;
				ChkLogExceptions.Enabled = false;
				TxtLogFile.Enabled = false;
				BtnLogFile.Enabled = false;


			}
		}

		private void BtnLogFile_Click(object sender, System.EventArgs e)
		{
			DlgSaveLog.ShowDialog();
		}

		private void DlgSaveLog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
		{
			TxtLogFile.Text = DlgSaveLog.FileName;
		}

		private void CboSite_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			Context.Site = (SiteCodeType) Enum.Parse(typeof(SiteCodeType), CboSite.Text);
			ErrorLanguageCodeType lang = eBay.Service.Util.ErrorLanguageUtility.GetDefaultErrorLanguageCodeType(Context.Site);
			CboLanguage.SelectedIndex = CboLanguage.Items.IndexOf(lang.ToString());

		
		}

		private void CboLanguage_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			Context.ErrorLanguage = (ErrorLanguageCodeType) Enum.Parse(typeof(ErrorLanguageCodeType), CboLanguage.Text);

		}		
		
		private void replaceConfigSettings() 
		{ 
			XmlDocument xDoc= new XmlDocument();
			AppDomain myDomain = System.AppDomain.CurrentDomain;
			string sFileName = myDomain.SetupInformation.ConfigurationFile;
       
			try 
			{ 
				AppSettingHelper.SetAppSetting(AppSettingHelper.DEV_ID, Context.ApiCredential.ApiAccount.Developer);				
				AppSettingHelper.SetAppSetting(AppSettingHelper.APP_ID, Context.ApiCredential.ApiAccount.Application);				
				AppSettingHelper.SetAppSetting(AppSettingHelper.CERT_ID, Context.ApiCredential.ApiAccount.Certificate);				
				//annotated by peter
//				AppSettingHelper.SetAppSetting(AppSettingHelper.EBAY_USER, Context.ApiCredential.eBayAccount.UserName);
//				AppSettingHelper.SetAppSetting(AppSettingHelper.EBAY_PASSWORD, Context.ApiCredential.eBayAccount.Password);
				AppSettingHelper.SetAppSetting(AppSettingHelper.API_TOKEN, Context.ApiCredential.eBayToken);
				AppSettingHelper.SetAppSetting(AppSettingHelper.API_SERVER_URL, Context.SoapApiServerUrl);				
				AppSettingHelper.SetAppSetting(AppSettingHelper.SIGNIN_URL, Context.SignInUrl);				
				AppSettingHelper.SetAppSetting(AppSettingHelper.EPS_SERVER_URL, Context.EPSServerUrl);				
				AppSettingHelper.SetAppSetting(AppSettingHelper.VERSION, Context.Version);
				AppSettingHelper.SetAppSetting(AppSettingHelper.TIME_OUT, Context.Timeout.ToString());				
				AppSettingHelper.SetAppSetting(AppSettingHelper.LOG_FILE_NAME, 
									((eBay.Service.Util.FileLogger)(Context.ApiLogManager.ApiLoggerList.ItemAt(0))).FileName);
				AppSettingHelper.SaveAppSettings();
			} 
			catch(Exception ex) 
			{ 
        #if DEBUG 
				System.Windows.Forms.MessageBox.Show("replaceConfigSettings()"+ex.Message); 
        #else 
          System.Windows.Forms.MessageBox.Show("Unable to save changes"); 
        #endif 
				xDoc = null; 
			}   

		}
					
	}
}
