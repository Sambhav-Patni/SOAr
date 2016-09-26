using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelBeautifier
{
    public class MonthlyReportSummaryGenerator
    {
        /// <summary>
        /// Genrates summary of APD Monthly Report excel sheets and stores them into an .xls file
        /// </summary>
        /// <param name="pathOfMonthlyReport"> Path of source monthly report</param>
        /// <param name="pathOfResolvedDataFile">Path of resolved.xlsx (Required for average time to resolve sheet) </param>
        /// <param name="storagePathOfOutput">Where should I store the output file? File will be in .xls format</param>
        public void ReadAndGenerateSummary(string pathOfMonthlyReport, string pathOfResolvedDataFile, string storagePathOfOutput)
        {
          
            //read input files            
            var workbookOfMonthlyReport = WorkbookFactory.Create(pathOfMonthlyReport);

            var listOfSummarizedExcelData = new List<ExcelSheetDto>();
            List<ISheet> sheetsToSummarize = new List<ISheet>();
           
            for (int i_it = 0; i_it < workbookOfMonthlyReport.NumberOfSheets; i_it++) {
                var sheet = workbookOfMonthlyReport.GetSheetAt(i_it);
                if (ExcelReportSummaryConfiguration.ListOfNamesOfSheetsToProcess.Contains(sheet.SheetName)) {
                    sheetsToSummarize.Add(sheet);
                }
            }
     
                        
            
            foreach (var sheetToSummarize in sheetsToSummarize) {
                ExcelSheetDto excelData = ApdMonthlyReportHelper.SanitizeAndTransform(sheetToSummarize);
                excelData.Name = sheetToSummarize.SheetName;
                excelData.Body = ReportDataSummarizer.Summarize(excelData.Body);
                listOfSummarizedExcelData.Add(excelData);
            }
            var workbookOfResolvedReport = WorkbookFactory.Create(pathOfResolvedDataFile);
            var resolvedSheet = workbookOfResolvedReport.GetSheet(ExcelReportSummaryConfiguration.NameOfResolvedSheet);

            //TODO: code AverageTimeToResolveSheet.ParseAndSummarize
            //DONE
            if (resolvedSheet != null) {
                ExcelSheetDto averageTimeToResolveSheet = AverageTimeToResolveSheet.ParseAndSummarize(resolvedSheet);
                averageTimeToResolveSheet.Name = "Average Time To Resolve";
                listOfSummarizedExcelData.Add(averageTimeToResolveSheet);
            }
            else {
                Console.WriteLine(" 'Average time to resolve' sheet not found");
            }
            //sort 
            listOfSummarizedExcelData.ForEach(sheet => {

                sheet.Body.Sort((a, b) => { 
                    int a_in = ExcelReportSummaryConfiguration.
                                ListOfNamesOfPrimaryTierOfProductCategory.
                                IndexOf(a.ListOfCells[0].GetValueAsString());
                    int b_in = ExcelReportSummaryConfiguration.
                                ListOfNamesOfPrimaryTierOfProductCategory.
                                IndexOf(b.ListOfCells[0].GetValueAsString());
                    if(a_in != b_in)
                        return a_in.CompareTo(b_in);

                    a_in = ExcelReportSummaryConfiguration.
                            ListOfNamesOfSecondaryTierOfProductCategoryApdClaims.
                               IndexOf(a.ListOfCells[1].GetValueAsString());
                    b_in = ExcelReportSummaryConfiguration.
                             ListOfNamesOfSecondaryTierOfProductCategoryApdClaims.
                               IndexOf(b.ListOfCells[1].GetValueAsString());
                    return a_in.CompareTo(b_in);
                });
            
            });

            //write output files
            IWorkbook outputWorkbook = ExcelWriter.ToExcelWorkbook(listOfSummarizedExcelData);
            
            AddStylesToWorkBook(outputWorkbook);

            var streamOfOutputFile = new FileStream(storagePathOfOutput, FileMode.Create, FileAccess.Write, FileShare.None);
            outputWorkbook.Write(streamOfOutputFile);

        }

        private void AddStylesToWorkBook(IWorkbook outputWorkbook)
        {
            //add styles to each sheet
            var styler = new ReportSheetStyler();
            for (int i = 0; i < outputWorkbook.NumberOfSheets; i++) {
                styler.StyleSheet( outputWorkbook.GetSheetAt(i));
            }
            styler.AddDocumentInformation((HSSFWorkbook)outputWorkbook);
            
           


        }
    }
}
