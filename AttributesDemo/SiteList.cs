using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using eBay.Service.Core.Soap;

namespace ItemSpecificsDemo
{
    public partial class SiteListForm : Form
    {
        private AttributesController controller = null;
        
        public SiteListForm(AttributesController controller)
        {
            this.controller = controller;

            InitializeComponent();
            InitVersion();
            InitSiteList();
        }

        private void InitSiteList()
        {
            this.siteListcomboBox.Items.Add(new ListItem("US", "US"));
            this.siteListcomboBox.Items.Add(new ListItem("eBayMotors", "eBayMotors"));
            this.siteListcomboBox.Items.Add(new ListItem("Canada", "Canada"));
            this.siteListcomboBox.Items.Add(new ListItem("UnitedKingdom", "UK"));
            this.siteListcomboBox.Items.Add(new ListItem("Australia", "Australia"));
            this.siteListcomboBox.Items.Add(new ListItem("Austria", "Austria"));
            this.siteListcomboBox.Items.Add(new ListItem("BelgiumFrench", "Belgium_French"));
            this.siteListcomboBox.Items.Add(new ListItem("France", "France"));
            this.siteListcomboBox.Items.Add(new ListItem("Germany", "Germany"));
            this.siteListcomboBox.Items.Add(new ListItem("Italy", "Italy"));
            this.siteListcomboBox.Items.Add(new ListItem("BelgiumDutch", "Belgium_Dutch"));
            this.siteListcomboBox.Items.Add(new ListItem("Netherlands", "Netherlands"));
            this.siteListcomboBox.Items.Add(new ListItem("Spain", "Spain"));
            this.siteListcomboBox.Items.Add(new ListItem("Switzerland", "Switzerland"));
            this.siteListcomboBox.Items.Add(new ListItem("Taiwan", "Taiwan"));
            this.siteListcomboBox.SelectedIndex = 0;
        }

        private void InitVersion()
        {
            versionLabel.Text += controller.ApiContext.Version;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SiteCodeType site = (SiteCodeType)Enum.Parse(typeof(SiteCodeType), ((ListItem)this.siteListcomboBox.SelectedItem).Value, true);
                this.controller.ApiContext.Site = site;

                this.Hide();

                this.controller.ShowPleaseWaitDialog();
                this.controller.InitSiteFacade();
                this.controller.HidePleaseWaitDialog();

                CategoryListForm categoryListForm = (CategoryListForm)controller.FormTable[AttributesController.CATEGORY_LIST_FORM];

                //check if category id is inputted
                String catId = this.catIdTextBox.Text;
                if (catId != null && catId.Length > 0)
                {
                    categoryListForm.CatId = catId;
                }
                else
                {
                    categoryListForm.CatId = null;
                }
                
                
                categoryListForm.InitCategoryList();

                categoryListForm.Show();
                categoryListForm.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

    }
}