using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using eBay.Service.Core.Soap;
using SHDocVw;

namespace ItemSpecificsDemo
{
    public partial class ItemSpecificsForm : Form
    {

        private const string CUSTOME_ITEM_SPECIFICS_NAME = "itemSpecificName_";
        private const string CUSTOME_ITEM_SPECIFICS_VALUE = "itemSpecificValue_";
        private const string EBAY_CUSTOM_ITEM_SPECIFICS_NAME = "eBayItemSpecificName_";
        private const string EBAY_CUSTOM_ITEM_SPECIFICS_VALUE = "eBayItemSpecificValue_";
        private const string EBAY_CUSTOM_ITEM_SPECIFICS_NAME_CACHE = "eBayItemSpecificNameCache_";

        private AttributesController controller = null;

        public ItemSpecificsForm(AttributesController controller)
        {
            this.controller = controller;

            InitializeComponent();

            this.registerWebBrowserEventHandler();
        }

        private void registerWebBrowserEventHandler()
        {
            webBrowser.Navigate("about:blank");
            SHDocVw.WebBrowser wb = (SHDocVw.WebBrowser)webBrowser.ActiveXInstance;
            wb.BeforeNavigate2 += new DWebBrowserEvents2_BeforeNavigate2EventHandler(WebBrowser_BeforeNavigate2);
        }

        //cancel the navigation, intercept the posted data 
        private void WebBrowser_BeforeNavigate2(object pDisp, ref object URL, ref object Flags,
                                        ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        {
            if (PostData != null)
            {
                Cancel = true;
                string postDataText = System.Text.Encoding.ASCII.GetString(PostData as byte[]);
                postDataText = postDataText.Replace("\0", "");
                NameValueCollection request = ParsePostString(postDataText);
                string action = request["action"];
                if (action == "post")
                {
                    Post(request);
                }
                else if (action == "back")
                {
                    this.Hide();

                    CategoryListForm categorySiteForm = (CategoryListForm)controller.FormTable[AttributesController.CATEGORY_LIST_FORM];

                    categorySiteForm.Show();
                    categorySiteForm.BringToFront();
                }
            }
        }

        private void Post(NameValueCollection request)
        {
            try
            {
                NameValueListTypeCollection nvCol = GetAllCustomItemSpecificsNameValue(request);
                this.controller.CategoryFacade.ItemSpecificsCache = nvCol;

                this.Hide();

                ReturnPolicyForm returnPolicyForm = controller.FormTable[AttributesController.RETURN_POLICY_FORM] as ReturnPolicyForm;

                returnPolicyForm.InitializeReturnPolicy();

                returnPolicyForm.Show();
                returnPolicyForm.BringToFront();
            }
            catch (Exception ex)
            {
                    MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private NameValueCollection ParsePostString(string s)
        {
            NameValueCollection nvc = new NameValueCollection();

            foreach (string vp in Regex.Split(s, "&"))
            {
                string[] singlePair = Regex.Split(vp, "=");
                nvc.Add(singlePair[0], singlePair[1]);
            }

            return nvc;
        }

        public void DisplayItemSpecifics(NameRecommendationTypeCollection col)
        {
            webBrowser.Document.OpenNew(true);

            AddHtmlHeaders();

            System.Text.StringBuilder str = new System.Text.StringBuilder();
            str.Append("</td></tr>");
            GetItemSpecificsHtml(ref str, col);
            webBrowser.Document.Write(str.ToString());

            AddHtmlTails();
        }

        private void AddHtmlHeaders()
        {
            webBrowser.Document.Write("<html>");
            webBrowser.Document.Write("<head>");
            webBrowser.Document.Write("<title>Attribute Info</title>");
            webBrowser.Document.Write("</head>");
            webBrowser.Document.Write("<body>");
            webBrowser.Document.Write("<script language=\"javascript\">");
            webBrowser.Document.Write("function onClick(id)");
            webBrowser.Document.Write("{");
            webBrowser.Document.Write("document.all('action').value = id;");
            webBrowser.Document.Write("document.forms['APIForm'].submit();");
            webBrowser.Document.Write("}\n");

            //CustomSpecific JS
            webBrowser.Document.Write(ItemSpecificsDemo.Properties.Resources.CustomSpecific);

            webBrowser.Document.Write("</script>");

            webBrowser.Document.Write("<form name=\"APIForm\" id=\"APIForm\" method=\"post\" action=\"\">");
            webBrowser.Document.Write("<table align=\"center\" border=\"0\"><tr><td><img src=\"http://pics.ebaystatic.com/aw/pics/logos/logoEbay_x45.gif\"></td></tr><tr><td>");


        }

        private void AddHtmlTails()
        {
            webBrowser.Document.Write("</td></tr>");
            CategoryFacade cf = this.controller.CategoryFacade;
            if (cf.ConditionEnabled == ConditionEnabledCodeType.Enabled ||
                cf.ConditionEnabled == ConditionEnabledCodeType.Required)
            {
                webBrowser.Document.Write("<tr><td>");
                webBrowser.Document.Write("<font color='red'>Note:</font><br>");
                webBrowser.Document.Write("This category is condition enabled, <br>");
                webBrowser.Document.Write("please ignore item condition setting(if any) on this form.<br>");
                webBrowser.Document.Write("You can set item condition on AddItem form later.");
                webBrowser.Document.Write("</td></tr>");
            }
            webBrowser.Document.Write("<tr><td align=\"center\">");
            string s = "<input type=\"button\" name=\"btSubmit0\" value=\"&nbsp;Back&nbsp;\" id=\"btSubmit0\"";
            s += "onclick=\"javascript:onClick('back')\"" + "/>";
            webBrowser.Document.Write(s);
            webBrowser.Document.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
            s = "<input type=\"button\" name=\"btSubmit1\" value=\"Continue\" id=\"btSubmit1\"";
            s += "onclick=\"javascript:onClick('post')\"" + "/>";
            webBrowser.Document.Write(s);
            webBrowser.Document.Write("</td></tr></table>");
            webBrowser.Document.Write("<input type=\"hidden\" name=\"action\" value=\"display\"/>");
            webBrowser.Document.Write("</form></body></html>");
        }

        private void GetItemSpecificsHtml(ref System.Text.StringBuilder str, NameRecommendationTypeCollection col)
        {
            str.Append("<tr><td>");
            //CustomItemSpecificGroup div
            str.Append("<span>");
            str.Append("<div id=\"CustomItemSpecificGroup\">");

            GeteBayRecommendedItemSpecificsHtml(ref str, col);
            
            str.Append("</div>");
            str.Append("</span>");
            //NewCustomItemSpecific div
            str.Append("<span>");
            str.Append("<div id=\"NewCustomItemSpecific\">");
            str.Append("</div>");
            str.Append("</span>");
            //SuggestedSectionLyr div
            str.Append("<span>");
            str.Append("<div id=\"tray\" style=\"visibility: visible;\">");
            str.Append("<div id=\"Addmore\">");
            str.Append("<b>Add more Specifics</b>");
            str.Append("</div>");
            str.Append("<div id=\"msg\" style=\"display: none; visibility: hidden;\"> If you want to add another item specific, please remove one of the existing specifics and add a new one. </div>");
            str.Append("<div id=\"SuggestedSectionLyr\" style=\"padding-top: 10px; visibility: visible;\">");
            str.Append("</div>");
            str.Append("</div>");
            str.Append("</span>");
            //AddCustomLnk tip
            str.Append("<span>");
            str.Append("<div>");
            str.Append("<span id=\"AddCustomLnk\" onclick=\"addNewSpecific();\">");
            str.Append("<a href=\"javascript:void(0);\">");
            str.Append("<img src=\"http://pics.qa.ebaystatic.com/aw/pics/buttons/btnOptionAdd.gif\" hspace=\"1\" border=\"0\" align=\"absmiddle\" />");
            str.Append("</a>");
            str.Append("<a href=\"javascript:void(0);\"><b>Add a custom detail</b></a>");
            str.Append("</span>");
            str.Append("</div>");
            str.Append("</span>");
            str.Append("</td></tr>");
        }

        private void GeteBayRecommendedItemSpecificsHtml(ref System.Text.StringBuilder str, NameRecommendationTypeCollection col)
        {
            if (col == null) return;
            int suffix = 0;

            foreach (NameRecommendationType nrt in col)
            {
                str.Append(string.Format("<span id=\"SpecificLayer_{0}\" style=\"margin-top: 8px\">", suffix));
                str.Append("<div>");
                str.Append(string.Format("<div id=\"TagName_{0}\" style=\"margin-top: 10px\">", suffix));
                str.Append("<b>"); str.Append(nrt.Name); str.Append("</b>");
                str.Append("</div>");
                str.Append("<div>");
                str.Append(string.Format("<input type=\"text\" id=\"eBayItemSpecificName_{0}\" name=\"eBayItemSpecificName_{0}\" style=\"width:100px;height:21px;font-size:10pt;\">", suffix));
                str.Append("<span style=\"width:18px;border:0px solid red;\">");
                str.Append(string.Format("<select id=\"eBayItemSpecificValue_{0}\" name=\"eBayItemSpecificValue_{0}\" style=\"margin-left:-100px;width:118px; background-color:#FFEEEE;\" onchange=\"optionSelect(this.name,this.value);\">", suffix));
                str.Append("<option value=\" \">Enter Your Own</option>");

                //generate value
                foreach (ValueRecommendationType vrt in nrt.ValueRecommendation)
                {
                    str.Append(string.Format("<option value=\"{0}\">", vrt.Value));
                    str.Append(vrt.Value);
                    str.Append("</option>");
                }

                str.Append("</select>");
                str.Append("</span>");
                str.Append(string.Format("<a href=\"javascript:void(0);\" onclick=\"remove(this.id);return false;\" id=\"Remove_{0}\" class=\"navigation\">", suffix));
                str.Append("remove");
                str.Append("</a>");
                str.Append("</div>");
                str.Append(string.Format("<input type=\"hidden\" id=\"{2}{0}\" name=\"{2}{0}\" value=\"{1}\"></input>", suffix, nrt.Name, EBAY_CUSTOM_ITEM_SPECIFICS_NAME_CACHE));
                str.Append("</div>");
                str.Append("</span>");

                suffix++;
            }
        }

        private NameValueListTypeCollection GetAllCustomItemSpecificsNameValue(NameValueCollection request)
        {
            NameValueListTypeCollection atts = new NameValueListTypeCollection();
            NameValueListType nvList;

            foreach (string key in request.AllKeys)
            {
                int index;
                index = key.IndexOf(EBAY_CUSTOM_ITEM_SPECIFICS_NAME);
                if (index >= 0 && request[key] != string.Empty)
                {
                    string suffix = key.Substring(index + EBAY_CUSTOM_ITEM_SPECIFICS_NAME.Length);
                    nvList = new NameValueListType();
                    nvList.Name = request[EBAY_CUSTOM_ITEM_SPECIFICS_NAME_CACHE + suffix];
                    eBay.Service.Core.Soap.StringCollection strCol = new eBay.Service.Core.Soap.StringCollection();
                    strCol.Add(request[key]);
                    nvList.Value = strCol;
                    atts.Add(nvList);
                    continue;
                }

                index = key.IndexOf(CUSTOME_ITEM_SPECIFICS_VALUE);
                if (index >= 0 && request[key] != string.Empty)
                {
                    string suffix = key.Substring(index + CUSTOME_ITEM_SPECIFICS_VALUE.Length);
                    nvList = new NameValueListType();
                    nvList.Name = request[CUSTOME_ITEM_SPECIFICS_NAME + suffix];
                    eBay.Service.Core.Soap.StringCollection strCol = new eBay.Service.Core.Soap.StringCollection();
                    strCol.Add(request[key]);
                    nvList.Value = strCol;
                    atts.Add(nvList);
                    continue;
                }

            }

            return atts;
        }
    }
}