using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Windows.Forms;
using eBay.Service.Core.Sdk;
using eBay.Service.Core.Soap;
using Samples.Helper.Cache;
using eBay.Service.Call;

namespace ItemSpecificsDemo
{
    public class AttributesController : ApplicationContext 
    {
        public static string SELECT_SITE_FORM = "SelectSite";
        public static string CATEGORY_LIST_FORM = "CategoryListForm";
        public static string ATTRIBUTES_INFO_FORM = "AttributesInfoForm";
        public static string ITEM_SPECIFICS_FORM = "ItemSpecificsForm";
        public static string RETURN_POLICY_FORM = "ReturnPolicyForm";
        public static string ADD_ITEM_FORM = "AddItem";

        private ApiContext apiContext = null;

        //for forms cache, form name -> form
        private Hashtable formTable = new Hashtable();
        private Form pleaseWaitDlg = new PleaseWaitDlg();

        private SiteFacade siteFacade = null;
        private CategoryFacade categoryFacate = null;

        //constructor
        public AttributesController(ApiContext apiContext)
        {
            this.apiContext = apiContext;
        }

        public void AddForm(string formName, Form form)
        {
            form.FormClosing += new FormClosingEventHandler(DemoFinish);
            formTable.Add(formName, form);
        }

        public void StartDemo()
        {
            Form startForm = (Form)formTable[SELECT_SITE_FORM];
            startForm.Show();
        }

        public void DemoFinish(object sender, EventArgs e)
        {
            ExitThread();
        }

        public void ShowPleaseWaitDialog()
        {
            pleaseWaitDlg.Show();
            pleaseWaitDlg.BringToFront();
            pleaseWaitDlg.Update();
        }

        public void HidePleaseWaitDialog()
        {
            pleaseWaitDlg.Hide();
        }

        public void InitSiteFacade()
        {
            this.siteFacade = new SiteFacade(this.ApiContext);
        }

        public void InitCategoryFacade(string catId)
        {
            this.categoryFacate = new CategoryFacade(catId, this.ApiContext, this.siteFacade);
        }

        public ApiContext ApiContext
        {
            get { return apiContext; }
            set { apiContext = value; }
        }

        public Hashtable FormTable
        {
            get { return formTable; }
        }

        public SiteFacade SiteFacade
        {
            get { return this.siteFacade; }
        }

        public CategoryFacade CategoryFacade
        {
            get { return this.categoryFacate; }
        }

    }
}
