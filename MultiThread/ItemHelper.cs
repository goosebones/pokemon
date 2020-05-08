#region Copyright
//	Copyright (c) 2013 eBay, Inc.
//	
//	This program is licensed under the terms of the eBay Common Development and
//	Distribution License (CDDL) Version 1.0 (the "License") and any subsequent  
//	version thereof released by eBay.  The then-current version of the License can be 
//	found at http://www.opensource.org/licenses/cddl1.php and in the eBaySDKLicense 
//	file that is under the eBay SDK ../docs directory
#endregion

using System;
using System.Diagnostics;
using System.Collections;
using System.Collections.Specialized;
using System.Threading;
using System.Configuration;
using eBay.Service.Core.Sdk;
using eBay.Service.Core.Soap;
using eBay.Service.Util;
namespace GetItemTester
{
	public class ItemHelper
	{

		#region pulich method

		public static ItemType BuildItem()
		{
			ItemType item =new ItemType();
			item.AutoPay = false;
			item.BuyItNowPrice = new AmountType();
			item.BuyItNowPrice.Value = Convert.ToDouble("10.0");
			item.BuyItNowPrice.currencyID = CurrencyCodeType.USD;
			item.Country = CountryCodeType.US;
			item.Currency = CurrencyCodeType.USD;
			item.Description = "SDK item description";
			item.ListingDuration = "Days_7";
			item.ListingType = ListingTypeCodeType.Chinese;
			item.Location = "San Jose, CA";
			item.PaymentMethods = new BuyerPaymentMethodCodeTypeCollection();
			item.PaymentMethods.Add(BuyerPaymentMethodCodeType.PaymentSeeDescription);
			item.PrimaryCategory = new CategoryType();
			item.PrimaryCategory.CategoryID = "2312";
			item.Quantity = 1;
			item.RegionID = "0";
			item.ReservePrice = new AmountType();
			item.ReservePrice.Value = Convert.ToDouble("5.0");
			item.ReservePrice.currencyID = CurrencyCodeType.USD;
			item.ShippingDetails = new ShippingDetailsType();
			item.StartPrice = new AmountType();
			item.StartPrice.Value = Convert.ToDouble("1.0");
			item.StartPrice.currencyID = CurrencyCodeType.USD;
			item.SubTitle = "sub title";
			item.Title = "SDK item title";
			item.SecondaryCategory = new CategoryType();
			item.SecondaryCategory.CategoryID = "2369";
			//set shipping details
			item.ShippingDetails=getShippingDetails();
			//handling time
			item.DispatchTimeMax =1;
			// Payment
			BuyerPaymentMethodCodeTypeCollection arrPaymentMethods =
				new BuyerPaymentMethodCodeTypeCollection();
			arrPaymentMethods.Add(BuyerPaymentMethodCodeType.PayPal);
			item.PaymentMethods = arrPaymentMethods;
			item.PayPalEmailAddress = "me@ebay.com";
			//set return policy
			item.ReturnPolicy=GetPolicyForUS();

			return item;
		}
		

		/// <summary>
		/// get a policy for us site only.
		/// </summary>
		/// <returns></returns>
		public static ReturnPolicyType GetPolicyForUS()
		{
			ReturnPolicyType policy=new ReturnPolicyType();
			policy.Refund="MoneyBack";
			policy.ReturnsWithinOption="Days_14";
			policy.ReturnsAcceptedOption="ReturnsAccepted";
			policy.ShippingCostPaidByOption="Buyer";

			return policy;
		}

		#endregion


		#region private method
		/// <summary>
		/// get shipping details
		/// </summary>
		/// <returns></returns>
		private static ShippingDetailsType getShippingDetails()
		{
			// Shipping details.
			ShippingDetailsType sd = new ShippingDetailsType();
			SalesTaxType salesTax = new SalesTaxType();
			salesTax.SalesTaxPercent = 0.0825F;
			salesTax.SalesTaxPercentSpecified = true;
			salesTax.SalesTaxState = "CA";
			sd.SalesTax = salesTax;
			sd.AllowPaymentEdit = false;
			sd.AllowPaymentEditSpecified = true;
			sd.ApplyShippingDiscount = true;
			sd.ApplyShippingDiscountSpecified = true;			
			sd.PaymentInstructions = "eBay .Net SDK test instruction.";
			sd.ShippingType = ShippingTypeCodeType.Flat;
			ShippingServiceOptionsType st1 = new ShippingServiceOptionsType();
			st1.ShippingService = ShippingServiceCodeType.USPSPriority.ToString();
			st1.ShippingServiceAdditionalCost = new AmountType();
			st1.ShippingServiceAdditionalCost.Value = 0.1;
			st1.ShippingServiceAdditionalCost.currencyID = CurrencyCodeType.USD;
			st1.ShippingServiceCost = new AmountType();
			st1.ShippingServiceCost.Value = 0.1;
			st1.ShippingServiceCost.currencyID = CurrencyCodeType.USD;

			st1.ShippingServicePriority = 1;
			st1.ShippingServicePrioritySpecified = true;
			st1.ShippingInsuranceCost = new AmountType();
			st1.ShippingInsuranceCost.Value = 0.1;
			st1.ShippingInsuranceCost.currencyID = CurrencyCodeType.USD;
			ShippingServiceOptionsType st2 = new ShippingServiceOptionsType();
			st2.ExpeditedService = true;
			st2.ExpeditedServiceSpecified = true;
			st2.ShippingService = ShippingServiceCodeType.USPSFirstClass.ToString();
			st2.ShippingServiceAdditionalCost = new AmountType();
			st2.ShippingServiceAdditionalCost.Value = 0.1;
			st2.ShippingServiceAdditionalCost.currencyID = CurrencyCodeType.USD;

			st2.ShippingServiceCost = new AmountType();
			st2.ShippingServiceCost.Value = 0.1;
			st2.ShippingServiceCost.currencyID = CurrencyCodeType.USD;
			st2.ShippingServicePriority = 2;
			st2.ShippingServicePrioritySpecified = true;
			st2.ShippingInsuranceCost = new AmountType();
			st2.ShippingInsuranceCost.Value = 0.1;
			st2.ShippingInsuranceCost.currencyID = CurrencyCodeType.USD;

			sd.ShippingServiceOptions = new ShippingServiceOptionsTypeCollection();
			sd.ShippingServiceOptions.Add(st1);
			sd.ShippingServiceOptions.Add(st2);
			return sd;
		}

		#endregion
	}//end class
}
