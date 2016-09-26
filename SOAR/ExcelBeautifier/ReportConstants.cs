using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public static class ReportConstants
    {
        public static List<string> Tier1Headings = new List<string>() { 
                "Legacy - APD",                
                "Miscellaneous - APD", 
                "Platform & Appraisal Solutions - APD",      
                "Repair Solutions - APD",
                "Claims Solutions - APD"            
            };

        public static List<string> ApdOpsTier2Headings = new List<string>() { 
                "Partial Loss",
                "WC-Total Loss",               
                "EIG",        
                "WC-Miscellaneous"
            };

        public static List<string> ReportTabsToModify = new List<string>() { 
                "Average time to resolve",
                "Age From Open Date",
                "By Incident Status",
                "By Product",
                "Leakage Rate",               
                "SCR Report"                                  
            };
        public static List<string> SpecialReportTabsToModify = new List<string>() {                   
                "Average time to resolve",
                "Leakage Rate",                            
                "SCR Report"                                  
            };


        public static string Tier1ColumnHeading = "Product Categorization Tier1";
        public static string Tier2ColumnHeading = "Product Categorization Tier2";

        public static string ApdClaimsTier1Name = "Claims Solutions - APD";
        public static string MiscellaneousApdString = "Miscellaneous - APD";
        public static string ApdClaimsWcMiscellaneousString = "WC-Miscellaneous";

        public static string GetApdClaimsTier2HeadingParent(string p)
        {
            if (ApdOpsTier2Headings.Contains(p))
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


        internal static string GetTier1HeadingParent(string p)
        {
            if (Tier1Headings.Contains(p))
                return p;
            return "Miscellaneous - APD";
        }
    }
}
