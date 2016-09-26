using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class ReportDataMerger
    {
        public ReportMergeOption MergerOption { get; set; }
        public Func<ExcelRange, ExcelRange, object> Aggregate { get; set; }

       
        public bool SumReportRows(ReportRow group, ReportRow single)
        {
            ReportRow result;
            if (IsMergable(group, single, out result) == false) {
                return false;
            }


            group.Tier1Name = result.Tier1Name;
            group.Tier2Name = result.Tier2Name;

            for (int it_col = 0; it_col <= group.DataRange.End.Column - group.DataRange.Start.Column; it_col++) {

                if (group.IsFormula(it_col, true) || single.IsFormula(it_col, true) ) {
                    continue;
                }

                group.DataRange[it_col, true].Value =  ReportDataMerger.SumExcelCell(
                                                                group.DataRange[it_col , true],
                                                                single.DataRange[it_col , true]
                                                            );
                group.RowsContained += single.RowsContained;

                single.RowsContained = 0;
                single.DataRange[it_col, true].Value = 0;
            }

            return true;

        }

        public static object SumExcelCell(ExcelRange a, ExcelRange b)
        {
            string val_a = a[a.Start.Row, a.Start.Column].GetValue<string>();            
            string val_b = b[b.Start.Row, b.Start.Column].GetValue<string>();
            double dbl_a, dbl_b;

            //dbl_a, dbl_b defaults to zero if it is not a number
            if (val_a == "" || val_a == null || double.TryParse(val_a, out dbl_a) == false) {
                dbl_a = 0;
            }

            if (val_b == "" || val_b == null || double.TryParse(val_b, out dbl_b) == false) {
                dbl_b = 0;
            }
                        
            return dbl_a + dbl_b;


        }

        private bool IsMergable(ReportRow reportRow1, ReportRow reportRow2, out ReportRow resultant)
        {

            bool result = false;
            resultant = new ReportRow() {
                Tier1Name = "",
                Tier2Name = ""
            };

            if(reportRow1.isSameAs(reportRow2))
            {
                return false;
            }

            

            string report1Parent;
            string report2Parent;

            switch (MergerOption) {
                case ReportMergeOption.AllIntraTier:
                    result = reportRow1.Tier1Name == reportRow2.Tier1Name;
                    resultant.Tier1Name = reportRow1.Tier1Name;
                    resultant.Tier2Name = "";
                    break;

                case ReportMergeOption.ApdMiscellaniousMerge:
                    report1Parent = ReportConstants.
                                GetTier1HeadingParent(reportRow1.Tier1Name);
                    report2Parent = ReportConstants.
                                GetTier1HeadingParent(reportRow2.Tier1Name);
                    result = (report1Parent == report2Parent);
                    resultant.Tier1Name = ReportConstants.MiscellaneousApdString;
                    resultant.Tier2Name = "";


                    break;

                case ReportMergeOption.ApdClaimsIntraTierMerge:
                    report1Parent = ReportConstants.
                               GetApdClaimsTier2HeadingParent(reportRow1.Tier2Name);
                    report2Parent = ReportConstants.
                               GetApdClaimsTier2HeadingParent(reportRow2.Tier2Name);
                    result = (report1Parent == report2Parent);

                    resultant.Tier1Name = ReportConstants.ApdClaimsTier1Name;
                    resultant.Tier2Name = report1Parent;

                    break;

                default:

                    break;
            }

            return result;
        }


        
    }

}
