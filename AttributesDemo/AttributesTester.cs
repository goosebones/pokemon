using System;
using System.Collections.Generic;
using System.Windows.Forms;
using eBay.Service.Core.Sdk;
using eBay.Service.Util;
using System.Configuration;

namespace ItemSpecificsDemo
{
    static class AttributesTester
    {

        public const string KEY_API_URL = "ApiServerUrl";
        public const string KEY_LOGFILE = "logfile";
        public const string KEY_APITOKEN = "ApiToken";
        public const string KEY_VIEWITEM_URL = "ViewItemUrl";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            ApiContext apiContext = GetApiContext();

            AttributesController controller = new AttributesController(apiContext);
            controller.AddForm(AttributesController.SELECT_SITE_FORM, new SiteListForm(controller));
            controller.AddForm(AttributesController.CATEGORY_LIST_FORM, new CategoryListForm(controller));            
            controller.AddForm(AttributesController.ITEM_SPECIFICS_FORM, new ItemSpecificsForm(controller));
            controller.AddForm(AttributesController.RETURN_POLICY_FORM, new ReturnPolicyForm(controller));
            controller.AddForm(AttributesController.ADD_ITEM_FORM, new AddItemForm(controller));
            controller.StartDemo();

            Application.Run(controller);
        }

        //get parameters from config file and create
        //ApiContext object
        static ApiContext GetApiContext()
        {
            ApiContext cxt = new ApiContext();

            // set api server address
            cxt.SoapApiServerUrl = ConfigurationManager.AppSettings[KEY_API_URL];


            // set token
            ApiCredential ac = new ApiCredential();
            string token = ConfigurationManager.AppSettings[KEY_APITOKEN];
            ac.eBayToken = token;
            cxt.ApiCredential = ac;

            // initialize log.
            ApiLogManager logManager = null;
            string logPath = ConfigurationManager.AppSettings[KEY_LOGFILE];
            if (logPath.Length > 0)
            {
                logManager = new ApiLogManager();

                logManager.EnableLogging = true;

                logManager.ApiLoggerList = new ApiLoggerCollection();
                ApiLogger log = new FileLogger(logPath, true, true, true);
                logManager.ApiLoggerList.Add(log);
            }
            cxt.ApiLogManager = logManager;

            return cxt;
        }

    }
}