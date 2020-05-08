using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using eBay.Service.Core.Sdk;
using eBay.Service.Core.Soap;
using Samples.Helper.Cache;
using eBay.Service.Call;

namespace ItemSpecificsDemo
{
    //facade for site level API, cache site level API
    public class SiteFacade
    {
        
        private ApiContext apiContext = null;
        
        //for site categories cache, two level map, site -> category id -> cateory
        private Hashtable siteCategoriesTable = new Hashtable();
        //for site categories cache, site -> merged category
        private Hashtable siteMergedCategoriesTable = new Hashtable();

        //for eBay details cache, site -> eBayDetails
        private Hashtable siteEBayDetailsTable = new Hashtable();

        private Hashtable siteCategoriesFeaturesTable = new Hashtable();
        private Hashtable siteFeatureDefaultTable = new Hashtable();
        private Hashtable siteFeatureDefinitionsTable = new Hashtable();

        private const int TABLE_SIZE = 30000;

        //constructor
        public SiteFacade(ApiContext apiContext)
        {
            this.apiContext = apiContext;
            SyncAllCategoriesFeatures();
        }

        //get eBay details response
        public GeteBayDetailsResponseType GetEbayDetails()
        {
            if (!siteEBayDetailsTable.ContainsKey(apiContext.Site))
            {
                DetailsDownloader downloader = new DetailsDownloader(apiContext);
                GeteBayDetailsResponseType response = downloader.GeteBayDetails();

                siteEBayDetailsTable.Add(apiContext.Site, response);
                return response;
            }
            else
            {
                return siteEBayDetailsTable[apiContext.Site] as GeteBayDetailsResponseType;
            }
        }

        private void SyncAllCategoriesFeatures()
        {
            if (!siteCategoriesFeaturesTable.ContainsKey(apiContext.Site))
            {
                FeaturesDownloader downloader = new FeaturesDownloader(apiContext);
                GetCategoryFeaturesResponseType resp = downloader.GetCategoryFeatures();
                CategoryFeatureTypeCollection cfCol = resp.Category;
                Hashtable cfsTable = new Hashtable(TABLE_SIZE);
                foreach (CategoryFeatureType cf in cfCol)
                {
                    cfsTable.Add(cf.CategoryID, cf);
                }
                siteCategoriesFeaturesTable.Add(apiContext.Site, cfsTable);
                siteFeatureDefaultTable.Add(apiContext.Site, resp.SiteDefaults);
                siteFeatureDefinitionsTable.Add(apiContext.Site, resp.FeatureDefinitions);
            }
        }

        //get all categories table
        public Hashtable GetAllCategoriesTable()
        {
            if (!siteCategoriesTable.ContainsKey(apiContext.Site))
            {
                Hashtable catsTable = new Hashtable(TABLE_SIZE);
                CategoriesDownloader downloader = new CategoriesDownloader(apiContext);
                CategoryTypeCollection catsCol = downloader.GetAllCategories();

                foreach (CategoryType cat in catsCol)
                {
                    catsTable.Add(cat.CategoryID, cat);
                }
                siteMergedCategoriesTable.Add(apiContext.Site, catsCol);                
                siteCategoriesTable.Add(apiContext.Site, catsTable);
                return catsTable;
            }
            else
            {
                return siteCategoriesTable[apiContext.Site] as Hashtable;
            }
        }

        /**
         * Get categories using GetCategory2CS and GetCategories calls,
         * and merge the categories
         * 
         */
        public CategoryTypeCollection GetAllMergedCategories()
        {
            if (!siteMergedCategoriesTable.ContainsKey(apiContext.Site))
            {
                GetAllCategoriesTable();
            }            
            return siteMergedCategoriesTable[apiContext.Site] as CategoryTypeCollection;            
        }

        public Hashtable SiteCategoriesFeaturesTable
        {
            get { return this.siteCategoriesFeaturesTable; }
        }

        public Hashtable SiteFeatureDefaultTable
        {
            get { return this.siteFeatureDefaultTable; }
        }

        public Hashtable SiteFeatureDefinitionsTable
        {
            get { return this.siteFeatureDefinitionsTable; }
        }
       
    }
}
