using System;
using System.Configuration;
using System.Collections.Generic;
using eBay.Service.Call;
using eBay.Service.Core.Sdk;
using eBay.Service.Core.Soap;
using eBay.Service.Util;
using System.IO;
using System.Text;

namespace ConsoleAddItem
{
    /// <summary>
    /// A simple item adding sample,
    /// show basic flow to list an item to eBay Site using eBay SDK.
    /// </summary>
    class Program
    {
        private static ApiContext apiContext = null;

        static void Main(string[] args)
        {

            try {
                Console.WriteLine("+++++++++++++++++++++++++++++++++++++++");
                Console.WriteLine("+ Welcome to eBay SDK for .Net Sample +");
                Console.WriteLine("+ - ConsoleAddItem                    +");
                Console.WriteLine("+++++++++++++++++++++++++++++++++++++++");

                //[Step 1] Initialize eBay ApiContext object
                ApiContext apiContext = GetApiContext();

                //[Step 2] Create a new ItemType object
                ItemType item = BuildItem();


                //[Step 3] Create Call object and execute the Call
                AddItemCall apiCall = new AddItemCall(apiContext);
                Console.WriteLine("Begin to call eBay API, please wait ...");
                FeeTypeCollection fees = apiCall.AddItem(item);
                Console.WriteLine("End to call eBay API, show call result ...");
                Console.WriteLine();

                //[Step 4] Handle the result returned
                Console.WriteLine("The item was listed successfully!");
                double listingFee = 0.0;
                foreach (FeeType fee in fees)
                {
                    if (fee.Name == "ListingFee")
                    {
                        listingFee = fee.Fee.Value;
                    }
                }
                Console.WriteLine(String.Format("Listing fee is: {0}", listingFee));
                Console.WriteLine(String.Format("Listed Item ID: {0}", item.ItemID));
            } 
            catch (Exception ex)
            {
                Console.WriteLine("Fail to list the item : " + ex.Message);
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to close the program.");
            Console.ReadKey();

        }

        /// <summary>
        /// Populate eBay SDK ApiContext object with data from application configuration file
        /// </summary>
        /// <returns>ApiContext object</returns>
        static ApiContext GetApiContext()
        {
            //apiContext is a singleton,
            //to avoid duplicate configuration reading
            if (apiContext != null)
            {
                return apiContext;
            }
            else
            {
                apiContext = new ApiContext();

                //set Api Server Url
                apiContext.SoapApiServerUrl = 
                    ConfigurationManager.AppSettings["Environment.ApiServerUrl"];
                //set Api Token to access eBay Api Server
                ApiCredential apiCredential = new ApiCredential();
                apiContext.EPSServerUrl = ConfigurationManager.AppSettings["Environment.EPSServerURL"];
                apiCredential.eBayToken = 
                    ConfigurationManager.AppSettings["UserAccount.ApiToken"];
                apiContext.ApiCredential = apiCredential;
                //set eBay Site target to US
                apiContext.Site = SiteCodeType.US;

                //set Api logging
                apiContext.ApiLogManager = new ApiLogManager();
                apiContext.ApiLogManager.ApiLoggerList.Add(
                    new FileLogger("listing_log.txt", true, true, true)
                    );
                apiContext.ApiLogManager.EnableLogging = true;


                return apiContext;
            }
        }

        /// <summary>
        /// Build a sample item
        /// </summary>
        /// <returns>ItemType object</returns>
        static ItemType BuildItem()
        {
            ItemType item = new ItemType();

            // item title
            item.Title = "Goudon EX";
            // item description
            item.Description = "Groudon EX";

            // listing type
            item.ListingType = ListingTypeCodeType.Chinese;
            // listing price
            item.Currency = CurrencyCodeType.USD;
            item.StartPrice = new AmountType();
            item.StartPrice.Value = 99.99;
            item.StartPrice.currencyID = CurrencyCodeType.USD;

            // listing duration
            item.ListingDuration = "Days_7";
            var startTime = new DateTime(2020, 5, 11, 0, 30, 0, DateTimeKind.Utc);
            item.ScheduleTime = startTime;

            // item location and country
            item.Location = "Rochester, New York";
            item.Country = CountryCodeType.US;

            // listing category, 
            CategoryType category = new CategoryType();
            category.CategoryID = "2611"; //CategoryID = 11104 (CookBooks) , Parent CategoryID=267(Books)
            item.PrimaryCategory = category;
             
            // item quality
            item.Quantity = 1;

            // item condition, New
            item.ConditionID = 3000;

            // item specifics
            item.ItemSpecifics = buildItemSpecifics();

            // picture
            var pics = new PictureDetailsType();
            var s = new StringCollection();
            pics.PictureURL = s;


            eBay.Service.EPS.eBayPictureService eps = new eBay.Service.EPS.eBayPictureService(GetApiContext());
            UploadSiteHostedPicturesRequestType req = new UploadSiteHostedPicturesRequestType();

            
            var path = new DirectoryInfo(@"C:\Program Files (x86)\eBay\eBay .NET SDK v1131 Release\Samples\C#\ConsoleAddItem\groudon");
            var files = path.GetFiles();
            var i = 0;
            foreach (var file in files) 
            {
                byte[] arr = File.ReadAllBytes(file.FullName);
                Base64BinaryType b = new Base64BinaryType();
                b.Value = arr;
                req.PictureName = file.FullName + i.ToString();
                req.PictureData = b;

                UploadSiteHostedPicturesResponseType res = eps.UpLoadSiteHostedPicture(req, file.FullName);
                s.Add(res.SiteHostedPictureDetails.FullURL);

                Console.WriteLine("Uploaded picture: " + i.ToString());
                i++;
            }
            


            /*
            var path = new DirectoryInfo(@"C:\Users\Gunther\Documents\GitHub\guntherkroth\pic");
            var files = path.GetFiles();
            var i = 0;
            foreach (var file in files)
            {
                s.Add("https://guntherkroth.com/pic/" + file.Name);
                i++;
            }
            */

            item.PictureDetails = pics;


            Console.WriteLine("Do you want to use Business policy profiles to list this item? y/n");
            String input = Console.ReadLine();
            if (input.ToLower().Equals("y"))
            {
                item.SellerProfiles = BuildSellerProfiles();
            }
            else
            {
                // payment methods
                item.PaymentMethods = new BuyerPaymentMethodCodeTypeCollection();
                item.PaymentMethods.AddRange(
                    new BuyerPaymentMethodCodeType[] { BuyerPaymentMethodCodeType.PayPal }
                    );
                // email is required if paypal is used as payment method
                item.PayPalEmailAddress = "goose.bones12@gmail.com";

                // handling time is required
                item.DispatchTimeMax = 1;
                // shipping details
                item.ShippingDetails = BuildShippingDetails();

                // return policy
                item.ReturnPolicy = new ReturnPolicyType();
                item.ReturnPolicy.ReturnsAcceptedOption = "ReturnsNotAccepted";
            }
            //item Start Price
            AmountType amount = new AmountType();
            amount.Value = 99.99;
            amount.currencyID = CurrencyCodeType.USD;
            item.StartPrice = amount;

            
            return item;
        }

        /// <summary>
        /// Build sample SellerProfile details
        /// </summary>
        /// <returns></returns>
        static SellerProfilesType BuildSellerProfiles()
        {
            /*
             * Beginning with release 763, some of the item fields from
             * the AddItem/ReviseItem/VerifyItem family of calls have been
             * moved to the Business Policies API. 
             * See http://developer.ebay.com/Devzone/business-policies/Concepts/BusinessPoliciesAPIGuide.html for more
             * 
             * This example uses profiles that were previously created using this api.
             */

            SellerProfilesType sellerProfile = new SellerProfilesType();
      
            Console.WriteLine("Enter Return policy profile Id:");            
            sellerProfile.SellerReturnProfile = new SellerReturnProfileType();
            sellerProfile.SellerReturnProfile.ReturnProfileID = Int64.Parse(Console.ReadLine());

            Console.WriteLine("Enter Shipping profile Id:");            
            sellerProfile.SellerShippingProfile = new SellerShippingProfileType();
            sellerProfile.SellerShippingProfile.ShippingProfileID = Int64.Parse(Console.ReadLine());

            Console.WriteLine("Enter Payment profile Id:");            
            sellerProfile.SellerPaymentProfile = new SellerPaymentProfileType();
            sellerProfile.SellerPaymentProfile.PaymentProfileID = Int64.Parse(Console.ReadLine());

            return sellerProfile;
        }

        /// <summary>
        /// Build sample shipping details
        /// </summary>
        /// <returns>ShippingDetailsType object</returns>
        static ShippingDetailsType BuildShippingDetails()
        {
            // Shipping details
            ShippingDetailsType sd = new ShippingDetailsType();

            sd.ApplyShippingDiscountSpecified = true;
            sd.ApplyShippingDiscount = false;
            sd.CalculatedShippingDiscount = null;

            sd.FlatShippingDiscount = null;

            sd.GlobalShipping = false;
            sd.GlobalShippingSpecified = false;

            sd.SellerExcludeShipToLocationsPreferenceSpecified = true;
            sd.SellerExcludeShipToLocationsPreference = false;


            // Shipping type and shipping service options
            sd.ShippingType = ShippingTypeCodeType.Flat;
            ShippingServiceOptionsType shippingOptions = new ShippingServiceOptionsType();
            shippingOptions.ShippingService = ShippingServiceCodeType.USPSFirstClass.ToString();
            shippingOptions.ExpeditedService = false;
            shippingOptions.ExpeditedServiceSpecified = true;
            shippingOptions.FreeShipping = false;
            shippingOptions.FreeShippingSpecified = false;
            shippingOptions.LocalPickup = false;
            shippingOptions.LocalPickupSpecified = false;

            var amount = new AmountType();
            amount.Value = 2.95;
            amount.currencyID = CurrencyCodeType.USD;
            shippingOptions.ShippingServiceCost = amount;

            shippingOptions.ShippingInsuranceCost = null;

            sd.ShippingServiceOptions = new ShippingServiceOptionsTypeCollection(
                new ShippingServiceOptionsType[] { shippingOptions }
                );

            return sd;
        }

        /// <summary>
        /// Build sample item specifics
        /// </summary>
        /// <returns>ItemSpecifics object</returns>
        static NameValueListTypeCollection buildItemSpecifics()
        {        	  
	        //create the content of item specifics
            NameValueListTypeCollection nvCollection = new NameValueListTypeCollection();
            
            NameValueListType nv1 = new NameValueListType();
            nv1.Name = "Set";
            StringCollection nv1Col = new StringCollection();
            String[] strArr1 = new string[] { "EX Crystal Guardians" };
            nv1Col.AddRange(strArr1);
            nv1.Value = nv1Col;
            
            NameValueListType nv2 = new NameValueListType();
            nv2.Name = "Rarity";
            StringCollection nv2Col = new StringCollection();
            String[] strArr2 = new string[] { "Ultra Rare" };
            nv2Col.AddRange(strArr2);
            nv2.Value = nv2Col;

            NameValueListType nv3 = new NameValueListType();
            nv3.Name = "Features";
            StringCollection nv3Col = new StringCollection();
            String[] strArr3 = new string[] { "Holo" };
            nv3Col.AddRange(strArr3);
            nv3.Value = nv3Col;

            NameValueListType nv4 = new NameValueListType();
            nv4.Name = "Featured Cards";
            StringCollection nv4Col = new StringCollection();
            String[] strArr4 = new string[] { "Groudon" };
            nv4Col.AddRange(strArr4);
            nv4.Value = nv4Col;

            NameValueListType nv5 = new NameValueListType();
            nv5.Name = "Quantity";
            StringCollection nv5Col = new StringCollection();
            String[] strArr5 = new string[] { "1" };
            nv5Col.AddRange(strArr5);
            nv5.Value = nv5Col;


            nvCollection.Add(nv1);
            nvCollection.Add(nv2);
            nvCollection.Add(nv3);
            nvCollection.Add(nv4);
            nvCollection.Add(nv5);
            return nvCollection;
         }
    }
}
