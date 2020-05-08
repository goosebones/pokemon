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
	/// Summary description for GetCategoriesForm.
	/// </summary>
	public class FrmGetFeedback : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.GroupBox GrpResult;
		private System.Windows.Forms.Button BtnGetFeedback;
		private System.Windows.Forms.Label LblUserId;
		private System.Windows.Forms.TextBox TxtUserId;
		private System.Windows.Forms.Label LblFeedbacks;
		private System.Windows.Forms.ListView LstComments;
		private System.Windows.Forms.ColumnHeader ClmUser;
		private System.Windows.Forms.ColumnHeader ClmScore;
		private System.Windows.Forms.ColumnHeader ClmItemId;
		private System.Windows.Forms.ColumnHeader ClmType;
		private System.Windows.Forms.ColumnHeader ClmRole;
		private System.Windows.Forms.ColumnHeader ClmTime;
		private System.Windows.Forms.ColumnHeader ClmComment;
		private System.Windows.Forms.TextBox TxtPositive;
		private System.Windows.Forms.TextBox TxtNegative;
		private System.Windows.Forms.TextBox TxtScore;
		private System.Windows.Forms.Label LblPositive;
		private System.Windows.Forms.Label LblNegative;
		private System.Windows.Forms.Label LblScore;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetFeedback()
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
			this.BtnGetFeedback = new System.Windows.Forms.Button();
			this.GrpResult = new System.Windows.Forms.GroupBox();
			this.LblScore = new System.Windows.Forms.Label();
			this.LblNegative = new System.Windows.Forms.Label();
			this.LblPositive = new System.Windows.Forms.Label();
			this.TxtScore = new System.Windows.Forms.TextBox();
			this.TxtNegative = new System.Windows.Forms.TextBox();
			this.TxtPositive = new System.Windows.Forms.TextBox();
			this.LblFeedbacks = new System.Windows.Forms.Label();
			this.LstComments = new System.Windows.Forms.ListView();
			this.ClmUser = new System.Windows.Forms.ColumnHeader();
			this.ClmScore = new System.Windows.Forms.ColumnHeader();
			this.ClmItemId = new System.Windows.Forms.ColumnHeader();
			this.ClmType = new System.Windows.Forms.ColumnHeader();
			this.ClmComment = new System.Windows.Forms.ColumnHeader();
			this.ClmRole = new System.Windows.Forms.ColumnHeader();
			this.ClmTime = new System.Windows.Forms.ColumnHeader();
			this.LblUserId = new System.Windows.Forms.Label();
			this.TxtUserId = new System.Windows.Forms.TextBox();
			this.GrpResult.SuspendLayout();
			this.SuspendLayout();
			// 
			// BtnGetFeedback
			// 
			this.BtnGetFeedback.Location = new System.Drawing.Point(240, 40);
			this.BtnGetFeedback.Name = "BtnGetFeedback";
			this.BtnGetFeedback.Size = new System.Drawing.Size(128, 23);
			this.BtnGetFeedback.TabIndex = 9;
			this.BtnGetFeedback.Text = "GetFeedback";
			this.BtnGetFeedback.Click += new System.EventHandler(this.BtnGetFeedback_Click);
			// 
			// GrpResult
			// 
			this.GrpResult.Controls.Add(this.LblScore);
			this.GrpResult.Controls.Add(this.LblNegative);
			this.GrpResult.Controls.Add(this.LblPositive);
			this.GrpResult.Controls.Add(this.TxtScore);
			this.GrpResult.Controls.Add(this.TxtNegative);
			this.GrpResult.Controls.Add(this.TxtPositive);
			this.GrpResult.Controls.Add(this.LblFeedbacks);
			this.GrpResult.Controls.Add(this.LstComments);
			this.GrpResult.Location = new System.Drawing.Point(8, 72);
			this.GrpResult.Name = "GrpResult";
			this.GrpResult.Size = new System.Drawing.Size(544, 336);
			this.GrpResult.TabIndex = 10;
			this.GrpResult.TabStop = false;
			this.GrpResult.Text = "Result";
			// 
			// LblScore
			// 
			this.LblScore.Location = new System.Drawing.Point(280, 24);
			this.LblScore.Name = "LblScore";
			this.LblScore.Size = new System.Drawing.Size(56, 23);
			this.LblScore.TabIndex = 78;
			this.LblScore.Text = "Score:";
			// 
			// LblNegative
			// 
			this.LblNegative.Location = new System.Drawing.Point(144, 24);
			this.LblNegative.Name = "LblNegative";
			this.LblNegative.Size = new System.Drawing.Size(56, 23);
			this.LblNegative.TabIndex = 77;
			this.LblNegative.Text = "Negative:";
			// 
			// LblPositive
			// 
			this.LblPositive.Location = new System.Drawing.Point(16, 24);
			this.LblPositive.Name = "LblPositive";
			this.LblPositive.Size = new System.Drawing.Size(48, 23);
			this.LblPositive.TabIndex = 75;
			this.LblPositive.Text = "Positive:";
			// 
			// TxtScore
			// 
			this.TxtScore.Location = new System.Drawing.Point(336, 24);
			this.TxtScore.Name = "TxtScore";
			this.TxtScore.ReadOnly = true;
			this.TxtScore.Size = new System.Drawing.Size(56, 20);
			this.TxtScore.TabIndex = 74;
			this.TxtScore.Text = "";
			// 
			// TxtNegative
			// 
			this.TxtNegative.Location = new System.Drawing.Point(200, 24);
			this.TxtNegative.Name = "TxtNegative";
			this.TxtNegative.ReadOnly = true;
			this.TxtNegative.Size = new System.Drawing.Size(56, 20);
			this.TxtNegative.TabIndex = 73;
			this.TxtNegative.Text = "";
			// 
			// TxtPositive
			// 
			this.TxtPositive.Location = new System.Drawing.Point(64, 24);
			this.TxtPositive.Name = "TxtPositive";
			this.TxtPositive.ReadOnly = true;
			this.TxtPositive.Size = new System.Drawing.Size(56, 20);
			this.TxtPositive.TabIndex = 71;
			this.TxtPositive.Text = "";
			// 
			// LblFeedbacks
			// 
			this.LblFeedbacks.Location = new System.Drawing.Point(16, 56);
			this.LblFeedbacks.Name = "LblFeedbacks";
			this.LblFeedbacks.Size = new System.Drawing.Size(112, 23);
			this.LblFeedbacks.TabIndex = 12;
			this.LblFeedbacks.Text = "Feedback:";
			// 
			// LstComments
			// 
			this.LstComments.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						  this.ClmUser,
																						  this.ClmScore,
																						  this.ClmItemId,
																						  this.ClmType,
																						  this.ClmComment,
																						  this.ClmRole,
																						  this.ClmTime});
			this.LstComments.GridLines = true;
			this.LstComments.Location = new System.Drawing.Point(16, 88);
			this.LstComments.Name = "LstComments";
			this.LstComments.Size = new System.Drawing.Size(520, 240);
			this.LstComments.TabIndex = 13;
			this.LstComments.View = System.Windows.Forms.View.Details;
			// 
			// ClmUser
			// 
			this.ClmUser.Text = "Commenting User";
			this.ClmUser.Width = 100;
			// 
			// ClmScore
			// 
			this.ClmScore.Text = "Score";
			this.ClmScore.Width = 56;
			// 
			// ClmItemId
			// 
			this.ClmItemId.Text = "Item Id";
			this.ClmItemId.Width = 61;
			// 
			// ClmType
			// 
			this.ClmType.Text = "Type";
			this.ClmType.Width = 53;
			// 
			// ClmComment
			// 
			this.ClmComment.Text = "Comment";
			this.ClmComment.Width = 86;
			// 
			// ClmRole
			// 
			this.ClmRole.Text = "Role";
			// 
			// ClmTime
			// 
			this.ClmTime.Text = "Time";
			this.ClmTime.Width = 95;
			// 
			// LblUserId
			// 
			this.LblUserId.Location = new System.Drawing.Point(176, 8);
			this.LblUserId.Name = "LblUserId";
			this.LblUserId.Size = new System.Drawing.Size(64, 23);
			this.LblUserId.TabIndex = 13;
			this.LblUserId.Text = "User Id:";
			// 
			// TxtUserId
			// 
			this.TxtUserId.Location = new System.Drawing.Point(240, 8);
			this.TxtUserId.Name = "TxtUserId";
			this.TxtUserId.Size = new System.Drawing.Size(104, 20);
			this.TxtUserId.TabIndex = 14;
			this.TxtUserId.Text = "";
			// 
			// FrmGetFeedback
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(560, 413);
			this.Controls.Add(this.TxtUserId);
			this.Controls.Add(this.LblUserId);
			this.Controls.Add(this.GrpResult);
			this.Controls.Add(this.BtnGetFeedback);
			this.Name = "FrmGetFeedback";
			this.Text = "GetFeedback";
			this.GrpResult.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void BtnGetFeedback_Click(object sender, System.EventArgs e)
		{
			try
			{
				TxtPositive.Text = "";
				TxtNegative.Text = "";
				TxtScore.Text = "";
				LstComments.Items.Clear();
	
				GetFeedbackCall apicall = new GetFeedbackCall(Context);
				apicall.DetailLevelList.Add(DetailLevelCodeType.ReturnAll);

				FeedbackDetailTypeCollection feedbacks = apicall.GetFeedback(TxtUserId.Text);

		
				if (apicall.FeedbackSummary != null)
				{
					TxtPositive.Text = apicall.FeedbackSummary.UniquePositiveFeedbackCount.ToString();
					TxtNegative.Text =apicall.FeedbackSummary.UniqueNegativeFeedbackCount.ToString();
					TxtScore.Text = apicall.FeedbackScore.ToString();
				}


				foreach (FeedbackDetailType feedback in feedbacks)
				{
					string[] listparams = new string[7];
					listparams[0] = feedback.CommentingUser;
					listparams[1] = feedback.CommentingUserScore.ToString();
					listparams[2] = feedback.ItemID;
					listparams[3] = feedback.CommentType.ToString();
					listparams[4] = feedback.CommentText;
					listparams[5] = feedback.Role.ToString();
					listparams[6] = feedback.CommentTime.ToString();

					ListViewItem vi = new ListViewItem(listparams);
					LstComments.Items.Add(vi);

				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

	}
	
}
