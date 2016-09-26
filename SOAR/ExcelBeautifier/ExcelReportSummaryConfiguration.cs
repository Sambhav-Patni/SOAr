using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
   static public class ExcelReportSummaryConfiguration
    {

        public static readonly string NameOfHeadingOfPrimaryTierColumn = "Product Categorization Tier1";
        public static readonly string NameOfHeadingOfSecondaryTierColumn = "Product Categorization Tier2";

        public static string NameOfPrimaryTierApdMiscellaneous = "Miscellaneous - APD";
        public static string NameOfPrimaryTierApdClaims = "Claims Solutions - APD";
        public static string NameOfSecondaryTierApdClaimsWcMiscellaneous = "WC-Miscellaneous";
        public static string NameOfSecondaryTierApdClaimsPartialLoss = "Partial Loss";
       

        public static List<string> ListOfNamesOfSheetsToProcess;

        //so obvious here, but not in middle of function where it was
        public static string NameOfResolvedSheet = "resolved";

        public  const string NameOfByProductSheet = "By Product";
        public  const string NameOfScrReportSheet = "SCR Report";
        public  const string NameOfAgeFromOpenDateSheet = "Age From Open Date";
        public  const string NameOfByIncidentStatusSheet = "By Incident Status";
        public  const string NameOfLeakageRateSheet = "Leakage Rate";
        public  const string NameOfAverageTimeToResolveSheet = "Average time to resolve";
        public const string NameOfAverageTimeToResolveSheetInOutput = "Average Time To Resolve";


        public static List<string> ListOfNamesOfPrimaryTierOfProductCategory;

        public static List<string> ListOfNamesOfSecondaryTierOfProductCategoryApdClaims;

        //following list is sorted as required, not alphabeticaly
        public static List<ProductCategory> SortedListOfProductCategoryOrder;


        static ExcelReportSummaryConfiguration()
        {
            ListOfNamesOfSheetsToProcess = new List<string>() {
                NameOfByProductSheet,
                NameOfScrReportSheet,
                //NameOfAverageTimeToResolveSheet,  //Can't calculate Average here due to lack of information
                NameOfAgeFromOpenDateSheet,
                NameOfByProductSheet,
                NameOfLeakageRateSheet
            };

            ListOfNamesOfPrimaryTierOfProductCategory = new List<string>() { 
                "Legacy - APD",          
                "Platform & Appraisal Solutions - APD",      
                "Repair Solutions - APD",
                NameOfPrimaryTierApdMiscellaneous,  
                NameOfPrimaryTierApdClaims  
            };

            ListOfNamesOfSecondaryTierOfProductCategoryApdClaims = new List<string>() { 
               NameOfSecondaryTierApdClaimsPartialLoss ,
               "WC-Total Loss",               
               "EIG",        
               NameOfSecondaryTierApdClaimsWcMiscellaneous
            };
            
            
            SortedListOfProductCategoryOrder = new List<ProductCategory>();

            //adding product categorization tier 1 names
            foreach (var PrimaryTierName in ListOfNamesOfPrimaryTierOfProductCategory) {
                SortedListOfProductCategoryOrder.Add(new ProductCategory() {
                    PrimaryTierName = PrimaryTierName,
                    SecondaryTierName = ""
                });
            }
            SortedListOfProductCategoryOrder.Remove(SortedListOfProductCategoryOrder.Last()); //last entry of APD Claims is invalid

            //adding Apd Claims tier names
            foreach (var SecondaryTierName in ListOfNamesOfSecondaryTierOfProductCategoryApdClaims) {
                SortedListOfProductCategoryOrder.Add(new ProductCategory() {
                    PrimaryTierName = NameOfPrimaryTierApdClaims,
                    SecondaryTierName = SecondaryTierName
                });
            }

        }
        public static string GetApdClaimsTier2HeadingParent(string p)
        {
            if (ListOfNamesOfSecondaryTierOfProductCategoryApdClaims.Contains(p))
                return p;

            if (p == "WC-Review" ||
                p == "WC-Assignment" ||
                p == "WC-Common Functions" ||
                p == "WC-Compliance Manager" ||
                p == "WC-Compliance Manager" ||
                p == "SIP Estimate" ||
                p == "SIP Assignment"
            ) {
                return "Partial Loss";
            }

            return "WC-Miscellaneous";
        }
        public static ProductCategory FindMasterProductCategory(this ResolvedReportRowData rr)
        {
            var prod_T = new ProductCategory() {
                PrimaryTierName = rr.PrimaryTier,
                SecondaryTierName = rr.SecondaryTier
            };

            prod_T = prod_T.FindMasterProductCategory();
            return prod_T;
        }

        //extension method so now ProductCategory can identify its parent
        public static ProductCategory FindMasterProductCategory(this ProductCategory slave)
        {            
            if (slave.PrimaryTierName == NameOfPrimaryTierApdClaims) {
                return FindMasterProductCategoryFromApdClaimsChild(slave);
            }
            return FindMasterOfGeneralProductCategoryChild(slave);
        }

        private static ProductCategory FindMasterOfGeneralProductCategoryChild( ProductCategory slave)
        {
            var parent = new ProductCategory();
            //Apd miscellanous logic here
            parent.SecondaryTierName = "";

            if (ListOfNamesOfPrimaryTierOfProductCategory.Contains(slave.PrimaryTierName)) {
                parent.PrimaryTierName = slave.PrimaryTierName;
            }
            else {
                parent.PrimaryTierName = NameOfPrimaryTierApdMiscellaneous;
            }
            return parent;
        }

        private static ProductCategory FindMasterProductCategoryFromApdClaimsChild(ProductCategory slave)
        {
            var parent = new ProductCategory();
            parent.PrimaryTierName = NameOfPrimaryTierApdClaims;
            //Apd Claims logic here

            parent.SecondaryTierName = GetApdClaimsTier2HeadingParent(slave.SecondaryTierName);
                        
            return parent;
        }
    }
}
