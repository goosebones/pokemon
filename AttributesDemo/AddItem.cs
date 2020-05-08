using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using eBay.Service.Core.Soap;
using eBay.Service.Core.Sdk;
using eBay.Service.Call;
using eBay.Service.Util;

namespace ItemSpecificsDemo
{
    public partial class AddItemForm : Form
    {
        private AttributesController controller;

        public AddItemForm(AttributesController controller)
        {
            InitializeComponent();

            this.controller = controller;
        }

        public void InitForm()
        {
            populateListingTypeComboBox();
            populateShippingServiceComboBox();
            populateShipToLocationTable();
            populatePaymentMethodTable();

            ConditionEnabledCodeType condition = this.controller.CategoryFacade.ConditionEnabled;
            if (condition == ConditionEnabledCodeType.Enabled ||
                condition == ConditionEnabledCodeType.Required)
            {
                this.conditionComboBox.Enabled = true;
                this.populateConditionComboBox();
            }
            else
            {
                this.conditionComboBox.Enabled = false;
            }

            if (this.controller.ApiContext.Site != SiteCodeType.eBayMotors)
            {
                this.eBayMotorGroupBox.Enabled = false;
            }
            else
            {
                this.eBayMotorGroupBox.Enabled = true;
            }
        }

        //populate payment methods
        private void populatePaymentMethodTable()
        {
            BuyerPaymentMethodCodeTypeCollection paymentMethods = this.controller.CategoryFacade.PaymentMethod;
            this.paymentMethodTableLayoutPanel.Controls.Clear();
            foreach (BuyerPaymentMethodCodeType payment in paymentMethods)
            {
                CheckBox cb = new CheckBox();
                cb.AutoSize = true;
                cb.Text = payment.ToString();
                cb.Name = payment.ToString();
                this.paymentMethodTableLayoutPanel.Controls.Add(cb);
            }
        }

        private void populateShipToLocationTable()
        {
            this.shipToLocationTableLayoutPanel.Controls.Clear();

            GeteBayDetailsResponseType eBayDetails = this.controller.SiteFacade.GetEbayDetails();
            if (eBayDetails != null)
            {
                ShippingLocationDetailsTypeCollection locations = eBayDetails.ShippingLocationDetails;

                if (locations != null)
                {
                    foreach (ShippingLocationDetailsType location in locations)
                    {
                        CheckBox cb = new CheckBox();
                        cb.AutoSize = true;
                        cb.Text = location.Description;
                        cb.Name = location.ShippingLocation;

                        this.shipToLocationTableLayoutPanel.Controls.Add(cb);
                    }
                }
            }
        }

        //populate shipping service list
        private void populateShippingServiceComboBox()
        {
            this.shippingServiceComboBox.Items.Clear();
            
            GeteBayDetailsResponseType eBayDetails = this.controller.SiteFacade.GetEbayDetails();
            if (eBayDetails != null)
            {

                if (this.controller.ApiContext.Site == SiteCodeType.eBayMotors)
                {
                    ListItem item = new ListItem();
                    item.Name = "None";
                    item.Value = "None";
                    this.shippingServiceComboBox.Items.Add(item);
                }
                
                
                ShippingServiceDetailsTypeCollection shippingDetails = eBayDetails.ShippingServiceDetails;

                if (shippingDetails != null)
                {
                    foreach (ShippingServiceDetailsType shippingDetail in shippingDetails)
                    {
                        if (shippingDetail.ServiceType.Contains(ShippingTypeCodeType.Flat)
                            && shippingDetail.ShippingServiceID < 5000)
                        {
                            ListItem item = new ListItem();
                            item.Name = shippingDetail.Description;
                            item.Value = shippingDetail.ShippingService.ToString();
                            this.shippingServiceComboBox.Items.Add(item);
                        }
                    }
                }

                if (this.shippingServiceComboBox.Items.Count > 0)
                {
                    this.shippingServiceComboBox.SelectedIndex = 0;
                }
            }
        }

        //populate listing duration list
        private void populateListingDurationComboBox()
        {
            Hashtable listingType2DurationMap = this.controller.CategoryFacade.ListingType2DurationMap;
            string listingType = (this.listingTypeComboBox.SelectedItem as ListItem).Value;
            StringCollection durations = (StringCollection)listingType2DurationMap[listingType];

            this.durationComboBox.Items.Clear();
            foreach (string duration in durations)
            {
                ListItem item = new ListItem(duration, duration);
                durationComboBox.Items.Add(item);
            }

            if (this.durationComboBox.Items.Count > 0)
            {
                this.durationComboBox.SelectedIndex = 0;
            }
        }

        private void populateConditionComboBox()
        {
            ConditionValuesType conditionValues = this.controller.CategoryFacade.ConditionValues;
            ConditionTypeCollection conditions = conditionValues.Condition;
            this.conditionComboBox.Items.Clear();

            foreach (ConditionType condition in conditions)
            {
                ListItem item = new ListItem(condition.DisplayName, condition.ID.ToString());
                this.conditionComboBox.Items.Add(item);
            }

            if (this.conditionComboBox.Items.Count > 0)
            {
                this.conditionComboBox.SelectedIndex = 0;
            }
        }

        //populate listing types list
        private void populateListingTypeComboBox()
        {
            ICollection listingTypes = this.controller.CategoryFacade.ListingType2DurationMap.Keys;

            this.listingTypeComboBox.Items.Clear();

            foreach (string listingType in listingTypes)
            {
                ListItem item = new ListItem(listingType, listingType);
                this.listingTypeComboBox.Items.Add(item);
            }

            if (this.listingTypeComboBox.Items.Count > 0)
            {
                this.listingTypeComboBox.SelectedIndex = 0;
            }
        }

        private void listingTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.populateListingDurationComboBox();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            addItem();
        }

        //list to eBay
        private void addItem()
        {

            try
            {
                //
                ApiContext apiContext = this.controller.ApiContext;
                AddItemCall api = new AddItemCall(apiContext);

                // Create the item
                ItemType item = new ItemType();

                item.Site = apiContext.Site;
                item.Country = CountryCodeType.US;

                item.ListingType = (ListingTypeCodeType)Enum.Parse(typeof(ListingTypeCodeType), (this.listingTypeComboBox.SelectedItem as ListItem).Value);
                if (item.ListingType.Equals(ListingTypeCodeType.LeadGeneration))
                {
                    item.ListingSubtype2 = ListingSubtypeCodeType.ClassifiedAd;
                }
                item.Title = this.titleTextBox.Text;
                item.Description = this.descriptionTextBox.Text;
                item.Currency = this.getCurrencyType(item.Site);

                if (this.startPriceTextBox.Text != string.Empty)
                {
                    item.StartPrice = new AmountType();
                    item.StartPrice.currencyID = item.Currency;
                    item.StartPrice.Value = Convert.ToDouble(this.startPriceTextBox.Text);
                }

                if (this.binTextBox.Text != string.Empty)
                {
                    item.BuyItNowPrice = new AmountType();
                    item.BuyItNowPrice.currencyID = item.Currency;
                    item.BuyItNowPrice.Value = Convert.ToDouble(this.binTextBox.Text);
                }

                item.Quantity = Int32.Parse(this.quantityTextBox.Text);

                item.Location = this.locationTextBox.Text;
                item.ListingDuration = (this.durationComboBox.SelectedItem as ListItem).Value;

                ListingEnhancementsCodeTypeCollection enhancements = new ListingEnhancementsCodeTypeCollection();
                enhancements.Add(ListingEnhancementsCodeType.BoldTitle);
                item.ListingEnhancement = enhancements;


                // Set Item Condtion
                ConditionEnabledCodeType condition = this.controller.CategoryFacade.ConditionEnabled;
                if (condition == ConditionEnabledCodeType.Enabled ||
                    condition == ConditionEnabledCodeType.Required)
                {
                    ListItem li = this.conditionComboBox.SelectedItem as ListItem;
                    int conditionId = int.Parse(li.Value);
                    item.ConditionID = conditionId;
                }

                // Set Attributes property.
                item.PrimaryCategory = new CategoryType();
                item.PrimaryCategory.CategoryID = this.controller.CategoryFacade.CategoryID;
                
                if (this.controller.CategoryFacade.ItemSpecificEnabled == ItemSpecificsEnabledCodeType.Enabled &&
                    this.controller.CategoryFacade.ItemSpecificsCache != null)
                {
                    item.ItemSpecifics = this.controller.CategoryFacade.ItemSpecificsCache;
                }

                // Motor
                if (this.controller.ApiContext.Site == SiteCodeType.eBayMotors &&
                    this.subTitleTextBox.Text.Length > 0)
                {
                    MotorAttributeHelper mh = new MotorAttributeHelper(item.AttributeSetArray[0]);
                    mh.Subtitle = this.subTitleTextBox.Text;
                    if (this.depositAmountTextBox.Text.Length > 0)
                        mh.DepositAmount = Decimal.Parse(this.depositAmountTextBox.Text);
                }

                if ((this.shippingServiceComboBox.SelectedItem as ListItem).Value != "None")
                {
                    //add shipping information
                    item.ShippingDetails = getShippingDetails((this.shippingServiceComboBox.SelectedItem as ListItem).Value);
                }

                //add handling time
                item.DispatchTimeMax = 1;

                SellerProfilesType sellerProfile = new SellerProfilesType();
                //add return policy
                if (this.controller.CategoryFacade.ReturnPolicyEnabled)
                {
                    if (this.controller.CategoryFacade.ReturnPolicyProfileCache != null)
                    {
                        sellerProfile.SellerReturnProfile = this.controller.CategoryFacade.ReturnPolicyProfileCache;
                    }
                    else if (this.controller.CategoryFacade.ReturnPolicyCache != null)
                    {
                        item.ReturnPolicy = this.controller.CategoryFacade.ReturnPolicyCache;
                    }
                }
                if (paymentProfileIdTextBox.Text != "" || paymentProfileNameTextBox.Text != "")
                {
                    sellerProfile.SellerPaymentProfile = new SellerPaymentProfileType();
                    if (this.paymentProfileIdTextBox.Text != "")
                    {
                        sellerProfile.SellerPaymentProfile.PaymentProfileID = Int64.Parse(paymentProfileIdTextBox.Text);
                        sellerProfile.SellerPaymentProfile.PaymentProfileIDSpecified = true;
                    }
                    sellerProfile.SellerPaymentProfile.PaymentProfileName = paymentProfileNameTextBox.Text;
                }
                if (shippingProfileIdTextBox.Text != "" || shippingProfileNameTextBox.Text != "")
                {
                    sellerProfile.SellerShippingProfile = new SellerShippingProfileType();
                    if (this.shippingProfileIdTextBox.Text != "")
                    {
                        sellerProfile.SellerShippingProfile.ShippingProfileID = Int64.Parse(shippingProfileIdTextBox.Text);
                        sellerProfile.SellerShippingProfile.ShippingProfileIDSpecified = true;
                    }
                    sellerProfile.SellerShippingProfile.ShippingProfileName = shippingProfileNameTextBox.Text;
                }

                //get shipping locations
                item.ShipToLocations = getShppingLocations();
                //set payments
                setPaymentMethods(item);
                // set seller profiles
                item.SellerProfiles = sellerProfile;

                FeeTypeCollection fees = api.AddItem(item);

                viewItemInfo(item);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void viewItemInfo(ItemType item)
        {
            System.Threading.Thread.Sleep(3000);
            string format = ConfigurationManager.AppSettings[AttributesTester.KEY_VIEWITEM_URL];
            string url = String.Format(format, item.ItemID);
            System.Diagnostics.Process.Start(url);
        }

        //set PaymentMethods
        private void setPaymentMethods(ItemType item)
        {
            BuyerPaymentMethodCodeTypeCollection paymentCol = new BuyerPaymentMethodCodeTypeCollection();
            foreach (CheckBox cb in this.paymentMethodTableLayoutPanel.Controls)
            {
                if (cb.Checked)
                {
                    BuyerPaymentMethodCodeType paymentCodeType = (BuyerPaymentMethodCodeType)Enum.Parse(typeof(BuyerPaymentMethodCodeType), cb.Name, false);
                    paymentCol.Add(paymentCodeType);
                    if (paymentCodeType == BuyerPaymentMethodCodeType.PayPal && this.paypalEmailTextBox.Text != string.Empty)
                    {
                        item.PayPalEmailAddress = this.paypalEmailTextBox.Text;
                    }
                }
            }

            item.PaymentMethods = paymentCol ;
        }


        //get shipping locations
        private StringCollection getShppingLocations()
        {
            StringCollection strCol = new StringCollection();

            foreach (CheckBox cb in this.shipToLocationTableLayoutPanel.Controls)
            {
                if (cb.Checked)
                {
                    strCol.Add(cb.Name);
                }
            }

            return strCol;
        }

        private ShippingDetailsType getShippingDetails(String shippingService)
        {
            // Shipping details.
            SiteCodeType site = this.controller.ApiContext.Site;
            ShippingDetailsType sd = new ShippingDetailsType();
            SalesTaxType salesTax = new SalesTaxType();
            salesTax.SalesTaxPercent = 0.0825f;
            salesTax.SalesTaxState = "CA";
            sd.ApplyShippingDiscount = true;
            AmountType at = new AmountType();
            at.Value = 0.1;
            at.currencyID = this.getCurrencyType(site);
            //sd.InsuranceFee = at;
            //sd.InsuranceOption = InsuranceOptionCodeType.Optional;
            sd.PaymentInstructions = "eBay DotNet SDK test instruction.";

            // Set calculated shipping.
            sd.ShippingType = ShippingTypeCodeType.Flat;
            //
            ShippingServiceOptionsType st1 = new ShippingServiceOptionsType();
            st1.ShippingService = shippingService;
            at = new AmountType();
            at.Value = 0.1;
            at.currencyID = this.getCurrencyType(site);
            st1.ShippingServiceAdditionalCost = at;
            at = new AmountType();
            at.Value = 0.1;
            at.currencyID = this.getCurrencyType(site);
            st1.ShippingServiceCost = at;
            st1.ShippingServicePriority = 1;
            at = new AmountType();
            at.Value = 0.1;
            at.currencyID = this.getCurrencyType(site);
            st1.ShippingInsuranceCost = at;

            sd.ShippingServiceOptions = new ShippingServiceOptionsTypeCollection(new ShippingServiceOptionsType[] { st1 });

            return sd;
        }

        private CurrencyCodeType getCurrencyType(SiteCodeType siteId)
        {
            CurrencyCodeType curr = CurrencyCodeType.USD;
            switch (siteId)
            {
                case SiteCodeType.US:
                case SiteCodeType.eBayMotors:
                    curr = CurrencyCodeType.USD;
                    break;
                case SiteCodeType.UK:
                    curr = CurrencyCodeType.GBP;
                    break;
                case SiteCodeType.Canada:
                    curr = CurrencyCodeType.CAD;
                    break;
                case SiteCodeType.Australia:
                    curr = CurrencyCodeType.AUD;
                    break;
                case SiteCodeType.Taiwan:
                    curr = CurrencyCodeType.TWD;
                    break;
                case SiteCodeType.Switzerland:
                    curr = CurrencyCodeType.CHF;
                    break;
                case SiteCodeType.Poland:
                    curr = CurrencyCodeType.PLN;
                    break;
                case SiteCodeType.India:
                    curr = CurrencyCodeType.INR;
                    break;
                case SiteCodeType.France:
                case SiteCodeType.Germany:
                case SiteCodeType.Italy:
                case SiteCodeType.Netherlands:
                case SiteCodeType.Belgium_Dutch:
                case SiteCodeType.Belgium_French:
                case SiteCodeType.Austria:
                    curr = CurrencyCodeType.EUR;
                    break;
                default:
                    curr = CurrencyCodeType.USD;
                    break;
            }

            return curr;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();

            ReturnPolicyForm returnPolicyForm = controller.FormTable[AttributesController.RETURN_POLICY_FORM] as ReturnPolicyForm;

            returnPolicyForm.Show();
            returnPolicyForm.BringToFront();
        }

    }
}