using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using eBay.Service.Core.Soap;
using System.Threading;

namespace ItemSpecificsDemo
{
    public partial class CategoryListForm : Form
    {
        private AttributesController controller = null;
        private SortedList sortedLeafCategories = null;

        private String catId = null;

        public String CatId
        {
            get { return catId; }
            set { catId = value; }
        }
	

        public CategoryListForm(AttributesController controller)
        {
            this.controller = controller;

            InitializeComponent();
        }

        public void InitCategoryList()
        {
            //show a please wait dialog for this lengthy process
            controller.ShowPleaseWaitDialog();

            sortedLeafCategories = GetSortedLeafCategories();

            populateCategoryList();
            showResult();

            controller.HidePleaseWaitDialog();

        }

        private void showResult()
        {
            if (categoryListBox.Items.Count == 0)
            {
                this.resultLabel.Text = "No Result Found.";
                this.resultLabel.ForeColor = Color.Red;
            }
            else
            {
                this.resultLabel.Text = "Found " + categoryListBox.Items.Count + " Categories";
                this.resultLabel.ForeColor = Color.Blue;
            }
        }

        private SortedList GetSortedLeafCategories()
        {
            SortedList catSL = new SortedList();

            CategoryTypeCollection catCol = controller.SiteFacade.GetAllMergedCategories();

            foreach (CategoryType cat in catCol)
            {
                //remove non-leaf category
                if (!cat.LeafCategory) continue;
                catSL.Add(int.Parse(cat.CategoryID), cat);
            }

            return catSL;
        }

        private Boolean isCategoryIdProvided()
        {
            return this.catId != null;
        }

        private void populateCategoryList()
        {
            this.categoryListBox.Items.Clear();

            Hashtable allCategories = controller.SiteFacade.GetAllCategoriesTable();
            Hashtable cfsTable = this.controller.SiteFacade.SiteCategoriesFeaturesTable[this.controller.ApiContext.Site] as Hashtable;
            SiteDefaultsType siteDefaults = this.controller.SiteFacade.SiteFeatureDefaultTable[this.controller.ApiContext.Site] as SiteDefaultsType;

            for (int i = 0; i < sortedLeafCategories.Count; i++)
            {
                CategoryType cat = (CategoryType)sortedLeafCategories.GetByIndex(i);
                string categoryId = cat.CategoryID;
                String catName = cat.CategoryName;
                //check if category id is provided on SiteList form, if provided,
                //just list the provided category
                if (this.isCategoryIdProvided() && categoryId != catId)
                {
                    continue;
                }

                // Walk up the category hierarchy to see if itemspecifics is enabled for this category
                ItemSpecificsEnabledCodeType itemSpecificsEnabled = (siteDefaults.ItemSpecificsEnabledSpecified == true) ?
                    siteDefaults.ItemSpecificsEnabled : ItemSpecificsEnabledCodeType.Disabled;                
                while (true)
                {
                    CategoryFeatureType cft = cfsTable[categoryId] as CategoryFeatureType;
                    CategoryType currentCat = (CategoryType)allCategories[categoryId];
                    if (cft != null && cft.ItemSpecificsEnabledSpecified == true) 
                    {
                        itemSpecificsEnabled = cft.ItemSpecificsEnabled;
                        break;
                    }
                    if (currentCat.CategoryLevelSpecified == true && currentCat.CategoryLevel == 1)
                    {
                        break;
                    }
                    categoryId = currentCat.CategoryParentID.ItemAt(0);
                }


                //ignore category which has no attributes or item specifics
                if (itemSpecificsEnabled.Equals(ItemSpecificsEnabledCodeType.Disabled))
                {
                    continue;
                }

                string name;
                string value;
                if (catName != null && catName.Length > 1)
                {
                    name = cat.CategoryName + "(" + cat.CategoryID + ")";
                    value = cat.CategoryID;
                }
                else
                {
                    name = cat.CharacteristicsSets[0].Name + "[" + cat.CategoryID + "]";
                    value = cat.CategoryID;
                }
                categoryListBox.Items.Add(new ListItem(name, value));
            }
        }

        
        private void button2_Click(object sender, EventArgs e)
        {
            SiteListForm selectSiteForm = (SiteListForm)controller.FormTable[AttributesController.SELECT_SITE_FORM];

            this.Hide();
            selectSiteForm.Show();
            selectSiteForm.BringToFront();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                ListItem selectedItem = this.categoryListBox.SelectedItem as ListItem;
                if (selectedItem == null)
                {
                    MessageBox.Show("Please select a category first!");
                    return;
                }

                this.Hide();

                this.controller.ShowPleaseWaitDialog();

                this.controller.InitCategoryFacade(selectedItem.Value);


                //for category with both attributes and item specifics enabled,
                //we only use item specifics
                if (this.controller.CategoryFacade.ItemSpecificEnabled == ItemSpecificsEnabledCodeType.Enabled)
                {
                    NameRecommendationTypeCollection itemSpecifics = this.controller.CategoryFacade.NameRecommendationsTypes;

                    ItemSpecificsForm itemSpecificsForm = controller.FormTable[AttributesController.ITEM_SPECIFICS_FORM] as ItemSpecificsForm;

                    itemSpecificsForm.DisplayItemSpecifics(itemSpecifics);

                    itemSpecificsForm.Show();
                    itemSpecificsForm.BringToFront();

                    this.controller.HidePleaseWaitDialog();

                    return;
                }

                ReturnPolicyForm returnPolicyForm = controller.FormTable[AttributesController.RETURN_POLICY_FORM] as ReturnPolicyForm;

                returnPolicyForm.InitializeReturnPolicy();

                returnPolicyForm.Show();
                returnPolicyForm.BringToFront();

                this.controller.HidePleaseWaitDialog();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           

        }

        private void AttributesOnlyRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                populateCategoryList();
                showResult();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ItemSpecificsOnlyRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                populateCategoryList();
                showResult();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BothRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                populateCategoryList();
                showResult();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AllRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                populateCategoryList();
                showResult();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}