using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using eBay.Service.Core.Soap;
using eBay.Service.Core.Sdk;
using eBay.Service.Call;
using eBay.Service.Util;

namespace SoapLibraryDemo
{
	/// <summary>
	/// Summary description for FrmGetTokenStatus.
	/// </summary>
	public class FrmGetTokenStatus : System.Windows.Forms.Form
	{
		public ApiContext Context;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.TextBox textBox2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox textBox3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textBox4;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmGetTokenStatus()
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
			this.button1 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.textBox2 = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.textBox3 = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.textBox4 = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(40, 16);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(104, 23);
			this.button1.TabIndex = 0;
			this.button1.Text = "GetTokenStatus";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.Location = new System.Drawing.Point(160, 16);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(104, 23);
			this.button2.TabIndex = 0;
			this.button2.Text = "RevokeToken";
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Status:";
			// 
			// textBox1
			// 
			this.textBox1.Location = new System.Drawing.Point(112, 72);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(160, 20);
			this.textBox1.TabIndex = 2;
			this.textBox1.Text = "";
			// 
			// textBox2
			// 
			this.textBox2.Location = new System.Drawing.Point(112, 112);
			this.textBox2.Name = "textBox2";
			this.textBox2.Size = new System.Drawing.Size(160, 20);
			this.textBox2.TabIndex = 2;
			this.textBox2.Text = "";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(16, 112);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(88, 16);
			this.label2.TabIndex = 1;
			this.label2.Text = "EIASToken:";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 152);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 16);
			this.label3.TabIndex = 1;
			this.label3.Text = "ExpirationTime:";
			// 
			// textBox3
			// 
			this.textBox3.Location = new System.Drawing.Point(112, 152);
			this.textBox3.Name = "textBox3";
			this.textBox3.Size = new System.Drawing.Size(160, 20);
			this.textBox3.TabIndex = 2;
			this.textBox3.Text = "";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(16, 192);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(88, 16);
			this.label4.TabIndex = 1;
			this.label4.Text = "RevocationTime:";
			// 
			// textBox4
			// 
			this.textBox4.Location = new System.Drawing.Point(112, 192);
			this.textBox4.Name = "textBox4";
			this.textBox4.Size = new System.Drawing.Size(160, 20);
			this.textBox4.TabIndex = 2;
			this.textBox4.Text = "";
			// 
			// FrmGetTokenStatus
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 238);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.textBox1,
																		  this.label1,
																		  this.button1,
																		  this.button2,
																		  this.textBox2,
																		  this.label2,
																		  this.label3,
																		  this.textBox3,
																		  this.label4,
																		  this.textBox4});
			this.Name = "FrmGetTokenStatus";
			this.Text = "FrmGetTokenStatus";
			this.ResumeLayout(false);

		}
		#endregion

		private void getTokenStatus()
		{
			try 
			{
                if (!validateApiAccount(Context))
                {
                    MessageBox.Show("Pleasse fill in Api Account first.");
                    return;
                }
                
                GetTokenStatusCall gtsc = new GetTokenStatusCall(Context);
				TokenStatusType tst = gtsc.GetTokenStatus();
				textBox1.Text = tst.Status.ToString();
				textBox2.Text = tst.EIASToken;
				textBox3.Text = tst.ExpirationTime.ToString();
				if (tst.Status == TokenStatusCodeType.RevokedByApp ||
					tst.Status == TokenStatusCodeType.RevokedByeBay ||
					tst.Status == TokenStatusCodeType.RevokedByUser) 
				{
					textBox4.Enabled = true;
					textBox4.Text = tst.RevocationTime.ToString();
				} 
				else 
				{
					textBox4.Enabled = false;
					textBox4.Text = "";
				}
			} 
			catch (Exception ex) 
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

		private void button1_Click(object sender, System.EventArgs e)
		{
			getTokenStatus();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			try
			{
                if (!validateApiAccount(Context))
                {
                    MessageBox.Show("Pleasse fill in Api Account first.");
                    return;
                }

				RevokeTokenCall rtc = new RevokeTokenCall(Context);
				rtc.RevokeToken(false);
				getTokenStatus();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}
}
