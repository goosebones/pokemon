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
    public partial class ReturnPolicyForm : Form
    {
        private AttributesController controller;

        public ReturnPolicyForm(AttributesController controller)
        {
            InitializeComponent();

            this.controller = controller;
        }

        public void InitializeReturnPolicy()
        {

            if (controller.CategoryFacade.ReturnPolicyEnabled)
            {
                this.returnPolicyLabel.Text = "Return Policy is Enabled for Category : " + controller.CategoryFacade.CategoryID;
                this.returnPolicyGroupBox.Enabled = true;
                populateReturnPolicy();
            }
            else
            {
                this.returnPolicyLabel.Text = "Return Policy is not Enabled for Category : " + controller.CategoryFacade.CategoryID;
                this.returnPolicyGroupBox.Enabled = false;
            }
        }

        private void populateReturnPolicy()
        {

            this.controller.ShowPleaseWaitDialog();
            GeteBayDetailsResponseType ebayDetails = controller.SiteFacade.GetEbayDetails();
            this.controller.HidePleaseWaitDialog();

            if (ebayDetails == null)
            {
                this.returnPolicyGroupBox.Enabled = false;
                return;
            }

            ReturnsAcceptedDetailsTypeCollection returnsAcceptedCol = ebayDetails.ReturnPolicyDetails.ReturnsAccepted;
            RefundDetailsTypeCollection refundCol = ebayDetails.ReturnPolicyDetails.Refund;
            ReturnsWithinDetailsTypeCollection returnsWithinCol = ebayDetails.ReturnPolicyDetails.ReturnsWithin;
            ShippingCostPaidByDetailsTypeCollection costPaidByCol = ebayDetails.ReturnPolicyDetails.ShippingCostPaidBy;

            if (returnsAcceptedCol != null && returnsAcceptedCol.Count > 0)
            {
                this.returnAcceptedComboBox.Items.Clear();
                foreach (ReturnsAcceptedDetailsType accept in returnsAcceptedCol)
                {
                    ListItem item = new ListItem();
                    item.Name = accept.Description;
                    item.Value = accept.ReturnsAcceptedOption;
                    returnAcceptedComboBox.Items.Add(item);
                    returnAcceptedComboBox.SelectedIndex = 0;
                }
            }
            else
            {
                this.returnAcceptedLabel.Enabled = false;
                this.returnAcceptedComboBox.Enabled = false;
            }

            if (refundCol != null && refundCol.Count > 0)
            {
                this.refundTypeComboBox.Items.Clear();
                foreach (RefundDetailsType refund in refundCol)
                {
                    ListItem item = new ListItem();
                    item.Name = refund.Description;
                    item.Value = refund.RefundOption;
                    refundTypeComboBox.Items.Add(item);
                    refundTypeComboBox.SelectedIndex = 0;
                }
            }
            else
            {
                this.refundTypeLabel.Enabled = false;
                this.refundTypeComboBox.Enabled = false;
            }

            if (returnsWithinCol != null && returnsWithinCol.Count > 0)
            {
                this.returnWithinComboBox.Items.Clear();
                foreach (ReturnsWithinDetailsType within in returnsWithinCol)
                {
                    ListItem item = new ListItem();
                    item.Name = within.Description;
                    item.Value = within.ReturnsWithinOption;
                    returnWithinComboBox.Items.Add(item);
                    returnWithinComboBox.SelectedIndex = 0;
                }
            }
            else
            {
                this.ReturnWithinLabel.Enabled = false;
                this.returnWithinComboBox.Enabled = false;
            }

            if (costPaidByCol != null && costPaidByCol.Count > 0)
            {
                this.shippingCostPaidByComboBox.Items.Clear();
                foreach (ShippingCostPaidByDetailsType cost in costPaidByCol)
                {
                    ListItem item = new ListItem();
                    item.Name = cost.Description;
                    item.Value = cost.ShippingCostPaidByOption;
                    shippingCostPaidByComboBox.Items.Add(item);
                    shippingCostPaidByComboBox.SelectedIndex = 0;
                }
            }
            else
            {
                this.shippingCostPaidByLabel.Enabled = false;
                this.shippingCostPaidByComboBox.Enabled = false;
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            
            if (this.controller.CategoryFacade.ItemSpecificEnabled == ItemSpecificsEnabledCodeType.Enabled)
            {
                ItemSpecificsForm itemSpecificsForm = controller.FormTable[AttributesController.ITEM_SPECIFICS_FORM] as ItemSpecificsForm;

                itemSpecificsForm.Show();
                itemSpecificsForm.BringToFront();
            }           
            else
            {
                CategoryListForm categoryListForm = (CategoryListForm)controller.FormTable[AttributesController.CATEGORY_LIST_FORM];

                categoryListForm.Show();
                categoryListForm.BringToFront();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.returnPolicyGroupBox.Enabled)
                {
                    if (this.returnPolicyProfileIdTextBox.Text != "" || this.returnPolicyProfileNameTextBox.Text != "")
                    {
                        SellerReturnProfileType returnProfile = new SellerReturnProfileType();
                        if (this.returnPolicyProfileIdTextBox.Text != "")
                        {
                            returnProfile.ReturnProfileID = Int64.Parse(this.returnPolicyProfileIdTextBox.Text);
                            returnProfile.ReturnProfileIDSpecified = true;
                        }
                        returnProfile.ReturnProfileName = this.returnPolicyProfileNameTextBox.Text;
                        this.controller.CategoryFacade.ReturnPolicyProfileCache = returnProfile;
                        this.controller.CategoryFacade.ReturnPolicyCache = null;
                    }
                    else
                    {
                        this.controller.CategoryFacade.ReturnPolicyCache = this.getReturnPolicy();
                        this.controller.CategoryFacade.ReturnPolicyProfileCache = null;
                    }
                }

                this.Hide();

                AddItemForm addItemForm = this.controller.FormTable[AttributesController.ADD_ITEM_FORM] as AddItemForm;
                controller.ShowPleaseWaitDialog();
                addItemForm.InitForm();
                controller.HidePleaseWaitDialog();
                addItemForm.Show();
                addItemForm.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private ReturnPolicyType getReturnPolicy()
        {
            ReturnPolicyType returnPolicy = new ReturnPolicyType();

            if (this.returnAcceptedComboBox.Enabled && this.returnAcceptedComboBox.SelectedItem != null)
            {
                returnPolicy.ReturnsAcceptedOption = (this.returnAcceptedComboBox.SelectedItem as ListItem).Value;
            }

            if (this.returnWithinComboBox.Enabled && this.returnWithinComboBox.SelectedItem != null)
            {
                returnPolicy.ReturnsWithinOption = (this.returnWithinComboBox.SelectedItem as ListItem).Value;
            }

            if (this.refundTypeComboBox.Enabled && this.refundTypeComboBox.SelectedItem != null)
            {
                returnPolicy.RefundOption = (this.refundTypeComboBox.SelectedItem as ListItem).Value;
            }

            if (this.shippingCostPaidByComboBox.Enabled && this.shippingCostPaidByComboBox.SelectedItem != null)
            {
                returnPolicy.ShippingCostPaidByOption = (this.shippingCostPaidByComboBox.SelectedItem as ListItem).Value;
            }

            returnPolicy.Description = this.descriptionTextBox.Text;

            return returnPolicy;

        }

       
    }
}