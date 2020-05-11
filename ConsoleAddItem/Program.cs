using System;
using System.Configuration;
using eBay.Service.Call;
using eBay.Service.Core.Sdk;
using eBay.Service.Core.Soap;
using eBay.Service.Util;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ListCards
{
    /// <summary>
    /// Create listings for cards in an external Excel Spreadsheet.
    /// </summary>
    class Program
    {
        private static ApiContext apiContext = null;

        static void Main(string[] args)
        {

            Console.WriteLine("+++++++++++++++++++++++++++++++++++++++");
            Console.WriteLine("+  Listing Individual Pokemon Cards   +");
            Console.WriteLine("+++++++++++++++++++++++++++++++++++++++\n");

            // Initialize eBay ApiContext object
            ApiContext apiContext = GetApiContext();

            // get reference to Spreadsheet
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workbook = xlApp.Workbooks.Open(@"C:\Users\Gunther\Desktop\pokemonSpread.xlsx");
            Excel.Worksheet sheet = workbook.Sheets[1];
            Excel.Range range = sheet.UsedRange;
            var rowCount = range.Rows.Count;
            var cells = range.Cells;

            // each row in the sheet is one card to list 
            for (int row = 2; row < rowCount + 1; row++)
            {
                try
                {
                    Console.WriteLine();

                    // cell contains a flag if this card was already listed
                    var listed = cells[row, 1].Value2;
                    if (listed == "Y")
                    {
                        continue;
                    }

                    // get details for the card 
                    var id = cells[row, 2].Value2.ToString();
                    var name = cells[row, 3].Value2;
                    var number = cells[row, 4].Value2;
                    var foil = cells[row, 5].Value2;
                    var rarity = cells[row, 6].Value2;
                    var set = cells[row, 7].Value2;
                    var condition = cells[row, 8].Value2;
                    var defects = cells[row, 9].Value2;
                    var location = cells[row, 10].Value2;
                    var price = cells[row, 11].Value2;

                    // build the eBay listing
                    Console.WriteLine("Listing Card #" + id);
                    var title = BuildItemTitle(name, number, foil, set, condition);
                    
                    // eBay titles cannot be longer that 80 characters
                    if (title.Length > 80)
                    {
                        Console.WriteLine("Did not list: " + name);
                        Console.WriteLine("Title too long");
                        // skip this card if its too long
                        // this card must be changed in the Spreadsheet,
                        // or it can be listed manually
                        continue;
                    }

                    ItemType item = BuildItem(id, title, name, foil, rarity, set, condition, defects, location, price);
                    
                    // Create Call object and execute the Call
                    Console.WriteLine("Calling API");
                    AddItemCall apiCall = new AddItemCall(apiContext);
                    FeeTypeCollection fees = apiCall.AddItem(item);
                    
                    // alert success and update Spreadsheet flag
                    Console.WriteLine("Listed Item");
                    sheet.Cells[row, 1] = "Y";
                    
                    // eBay api call will return any fees associated with a listing
                    double listingfee = 0.0;
                    foreach (FeeType fee in fees)
                    {
                        if (fee.Name == "ListingFee")
                        {
                            listingfee = fee.Fee.Value;
                        }
                    }
                    Console.WriteLine("Fees: " + listingfee);
                    Console.WriteLine("ItemID: " + item.ItemID);

                    // if there are listing fees, stop 
                    if (listingfee > 0.0)
                    {
                        Console.WriteLine("\n\nStopping. Listing fees accumulated.");
                        Console.ReadKey();
                        Environment.Exit(0);
                    }
                    
                } catch (Exception ex)
                {
                    // some error in listing item
                    // this will most likely be thrown by the eBay api call
                    Console.WriteLine("Failed to list the item: " + ex.Message);
                }
            }

            // clean up the Excel Spreadsheet
            workbook.Save();
            xlApp.DisplayAlerts = false;
            workbook.Close(false);
            xlApp.Quit();
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(xlApp);

            // alert success
            Console.WriteLine();
            Console.WriteLine("Finished.");
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
        /// Build the title of an Item listing
        /// </summary>
        /// <param name="name">Name of card</param>
        /// <param name="number">Set number</param>
        /// <param name="foil">Foil of card</param>
        /// <param name="set">Set the card belongs to</param>
        /// <param name="condition">Condition of card</param>
        /// <returns>Formatted title for the card listing</returns>
        static string BuildItemTitle(string name, string number, string foil, string set, string condition)
        {
            Console.WriteLine("Building title");

            var title = name + " " + number + " " + foil + " " + set + " ";

            switch (condition)
            {
                case "M":
                    title += "NM/M Mint";
                    break;
                case "NM":
                    title += "NM/M Near Mint";
                    break;
                case "LP":
                    title += "LP Lightly Played";
                    break;
                case "MP":
                    title += "MP Moderately Played";
                    break;
                case "HP":
                    title += "HP Heavily Played";
                    break;
                case "D":
                    title += "Damaged";
                    break;
                default:
                    break;
            }
            
            // check to see if 'Pokemon Card' can be added to title
            // title can be a max length of 80 characters
            if (title.Length <= 67)
            {
                title += " Pokemon Card";
            } else if (title.Length <= 72)
            {
                title += " Pokemon";
            } else if (title.Length <= 75)
            {
                title += " Card";
            }

            return title;
        }

        /// <summary>
        /// Build the description of a card listing.
        /// </summary>
        /// <param name="title">Title of card listing</param>
        /// <param name="condition">Condition of the card</param>
        /// <param name="defects">Any defects in the card</param>
        /// <param name="location">Location of defects</param>
        /// <returns></returns>
        static string BuildItemDescription(string title, string condition, string defects, string location)
        {
            // eBay uses HTML to format descriptions
            var description = "<div vocab=\"https://schema.org/\" typeof=\"Product\"><span property=\"description\">";
            description += title + "<br><br>";

            description += "Condition is ";
            switch (condition)
            {
                case "M":
                    description += "Mint";
                    break;
                case "NM":
                    description += "Near Mint";
                    break;
                case "LP":
                    description += "Lightly Played";
                    break;
                case "MP":
                    description += "Moderately Played";
                    break;
                case "HP":
                    description += "Heavily Played";
                    break;
                case "D":
                    description += "Damaged";
                    break;
                default:
                    break;
            }
            description += "." + "<br>";
            description += "Any card flaws/blemishes are visible in the pictures.<br>";
            description += "Please note " + defects + " on the " + location + " of the card.<br><br>";

            description += "All cards will be shipped with a KMC Perfect Fit sleeve, top loader, and bubble mailer. Combined shipping is also available.<br>";

            description += "Check out my other listings for more 2000s era cards (EX, Diamond & Pearl, Platinum).";

            description += "</span></div>";
            return description;
        }

        /// <summary>
        /// Build a card Item
        /// </summary>
        /// <param name="id">ID of card in Excel Spreadsheet</param>
        /// <param name="title">Title of card listing</param>
        /// <param name="name">name of card</param>
        /// <param name="foil">foil of card</param>
        /// <param name="rarity">rarity of card</param>
        /// <param name="set">set that the card belongs to</param>
        /// <param name="condition">condition of the card</param>
        /// <param name="defects">any defects of the card</param>
        /// <param name="location">location of the defects</param>
        /// <param name="price">starting price of the card</param>
        /// <returns></returns>
        static ItemType BuildItem(string id, string title, string name, string foil, string rarity, string set, string condition, string defects, string location, double price)
        {
            Console.WriteLine("Building Item");
            ItemType item = new ItemType();

            // item title
            item.Title = title;
            // item description
            item.Description = BuildItemDescription(title, condition, defects, location);

            // listing type
            item.ListingType = ListingTypeCodeType.Chinese;

            // listing price
            item.Currency = CurrencyCodeType.USD;
            item.StartPrice = new AmountType();
            item.StartPrice.Value = price;
            item.StartPrice.currencyID = CurrencyCodeType.USD;

            // listing duration
            item.ListingDuration = "Days_7";
            var startTime = new DateTime(2020, 5, 11, 2, 30, 0, DateTimeKind.Utc);
            item.ScheduleTime = startTime;

            // item location and country
            item.Location = "Rochester, New York";
            item.Country = CountryCodeType.US;

            // listing category, 
            CategoryType category = new CategoryType();
            category.CategoryID = "2611"; //CategoryID = 2611 
            item.PrimaryCategory = category;
             
            // item quality
            item.Quantity = 1;

            // item condition, Used
            item.ConditionID = 3000;

            // item specifics
            item.ItemSpecifics = buildItemSpecifics(set, rarity, foil, name);

            // upload pictures to EPS (eBay Picture Services)
            Console.Write("Uploading Pictures: ");

            var pics = new PictureDetailsType();
            var s = new StringCollection();
            pics.PictureURL = s;
            eBay.Service.EPS.eBayPictureService eps = new eBay.Service.EPS.eBayPictureService(GetApiContext());
            UploadSiteHostedPicturesRequestType req = new UploadSiteHostedPicturesRequestType();

            // pictures for each card are located in a folder that matches 
            // the ID from the Spreadsheet
            var folder = @"C:\Users\Gunther\Desktop\pics\";
            folder += id;

            // each picture in the folder gets uploaded
            var path = new DirectoryInfo(folder);
            var files = path.GetFiles();
            var i = 1;
            foreach (var file in files) 
            {
                // picstures are based an binary objects to EPS
                byte[] arr = File.ReadAllBytes(file.FullName);
                Base64BinaryType b = new Base64BinaryType();
                b.Value = arr;
                req.PictureName = file.FullName + i.ToString();
                req.PictureData = b;

                // api responds contains the URL of the picure 
                UploadSiteHostedPicturesResponseType res = eps.UpLoadSiteHostedPicture(req, file.FullName);
                s.Add(res.SiteHostedPictureDetails.FullURL);

                Console.Write(i.ToString() + " ");
                i++;
            }
            Console.WriteLine("done");

            // a collection of EPS urls are added the the Item listing
            item.PictureDetails = pics;

            // payment methods
            item.PaymentMethods = new BuyerPaymentMethodCodeTypeCollection();
            item.PaymentMethods.AddRange(
                new BuyerPaymentMethodCodeType[] { BuyerPaymentMethodCodeType.PayPal }
                );
            // email is required if paypal is used as payment method
            item.PayPalEmailAddress = "goose.bones12@gmail.com";

            // handling time is required
            item.DispatchTimeMax = 2;

            // shipping details
            item.ShippingDetails = BuildShippingDetails();

            // return policy
            item.ReturnPolicy = new ReturnPolicyType();
            item.ReturnPolicy.ReturnsAcceptedOption = "ReturnsNotAccepted";
            
            return item;
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
        /// Build Item Specifics for a card Item
        /// </summary>
        /// <param name="set">Set that the card belongs to</param>
        /// <param name="rarity">Rarity of the card</param>
        /// <param name="features">Foil of the card</param>
        /// <param name="name">Name of the card</param>
        /// <returns></returns>
        static NameValueListTypeCollection buildItemSpecifics(string set, string rarity, string features, string name)
        {        	  
	        //create the content of item specifics
            NameValueListTypeCollection nvCollection = new NameValueListTypeCollection();
            
            NameValueListType nv1 = new NameValueListType();
            nv1.Name = "Set";
            StringCollection nv1Col = new StringCollection();
            String[] strArr1 = new string[] { set };
            nv1Col.AddRange(strArr1);
            nv1.Value = nv1Col;
            
            NameValueListType nv2 = new NameValueListType();
            nv2.Name = "Rarity";
            StringCollection nv2Col = new StringCollection();
            String[] strArr2 = new string[] { rarity };
            nv2Col.AddRange(strArr2);
            nv2.Value = nv2Col;

            NameValueListType nv3 = new NameValueListType();
            nv3.Name = "Features";
            StringCollection nv3Col = new StringCollection();
            String[] strArr3 = new string[] { features };
            nv3Col.AddRange(strArr3);
            nv3.Value = nv3Col;

            NameValueListType nv4 = new NameValueListType();
            nv4.Name = "Featured Cards";
            StringCollection nv4Col = new StringCollection();
            String[] strArr4 = new string[] { name };
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
