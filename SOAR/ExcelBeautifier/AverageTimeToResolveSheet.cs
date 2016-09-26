using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelBeautifier
{
    public class AverageTimeToResolveSheet
    {
        internal static ExcelSheetDto ParseAndSummarize(NPOI.SS.UserModel.ISheet resolvedSheet)
        {
            var bodyOfAverageSheet = GetBodyOfAverageTimeToResolveSheet(resolvedSheet);

            var headersOfAverageSheet = GetHeadersOfAverageTimeToResolveSheet();

            ExcelSheetDto averageSheet = new ExcelSheetDto();
            averageSheet.Headers =  headersOfAverageSheet;
            averageSheet.Body = bodyOfAverageSheet;           

            return averageSheet;
        }

        private static List<ExcelRowDto> GetBodyOfAverageTimeToResolveSheet(NPOI.SS.UserModel.ISheet resolvedSheet)
        {
            var result = ToResolvedReportData(resolvedSheet).ToList();

            ///identify columns and transform data to excellRowDto
            //Prepare meta data and containers
            List<ProductCategory> sheetData = new List<ProductCategory>();

            foreach (var reportData in result) {
                updateAverageSheet(sheetData, reportData);
            }



            //translate columnal data to ExcelSheet
            var bodyOfAverageSheet = new List<ExcelRowDto>();
            foreach (var item in sheetData) {
                ExcelRowDto row_T = new ExcelRowDto();

                row_T.ListOfCells.Add(new ExcelCellDto(item.PrimaryTierName));
                row_T.ListOfCells.Add(new ExcelCellDto(item.SecondaryTierName));
                row_T.ListOfCells.AddRange(item.Properties);

                bodyOfAverageSheet.Add(row_T);
            }
            return bodyOfAverageSheet;
        }

        private static List<ExcelRowDto> GetHeadersOfAverageTimeToResolveSheet()
        {
            List<string> columnHeadingsOfAverageSheet = new List<string>() {
                ExcelReportSummaryConfiguration.NameOfHeadingOfPrimaryTierColumn,
                ExcelReportSummaryConfiguration.NameOfHeadingOfSecondaryTierColumn,
                "P1 Average Time(Days)",
                "P2 Average Time(Days)",                
                "P3 Average Time(Days)",
                "P4 Average Time(Days)",
                "P5 Average Time(Days)",
                "Blank Average Time(Days)",
                "Total Issues Resolved"
            };

            var excellCells = columnHeadingsOfAverageSheet.Select(head => new ExcelCellDto(head)).ToList();
            var headersOfAverageSheet = new List<ExcelRowDto>() { 
                 new ExcelRowDto () { ListOfCells = excellCells }
            };
            return headersOfAverageSheet;
        }

        /// <summary>
        /// It adds reportData to sheetData.
        /// Function has been extracted to simplify code not for reuse.
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="reportData"></param>
        private static void updateAverageSheet(List<ProductCategory> sheetData, ResolvedReportRowData reportData)
        {
            int numberOfColumns = 7;

            ProductCategory rowInsheet = sheetData.Find(
                r => reportData.PrimaryTier == r.PrimaryTierName &&
                        r.SecondaryTierName == reportData.SecondaryTier
            );

            if (rowInsheet == null) {
                rowInsheet = GetDefaultRowInSheet(reportData, numberOfColumns);
                sheetData.Add(rowInsheet);
            }

            int columnIndex = reportData.Priority - 1;
            //in case of blank priority, reportData.priority will be 0 and columnIndex -1
            if (columnIndex < 0) {
                columnIndex = 5;// this is the column number of blank priority
            }
            rowInsheet.Properties[columnIndex].SetValue(reportData.timeTaken.TotalDays / reportData.Count);
            rowInsheet.Properties[numberOfColumns - 1].AddValue(1);

            
        }

        private static IEnumerable<ResolvedReportRowData> ToResolvedReportData(NPOI.SS.UserModel.ISheet resolvedSheet)
        {

            var Data_T = ParseResolvedSheet(resolvedSheet);

            Data_T.ForEach(rr => {
                var prod_T = rr.FindMasterProductCategory();
                rr.PrimaryTier = prod_T.PrimaryTierName;
                rr.SecondaryTier = prod_T.SecondaryTierName;
            });

            var prodWithPriority_T = Data_T.GroupBy(rr => new ResolvedReportRowData {
                PrimaryTier = rr.PrimaryTier,
                SecondaryTier = rr.SecondaryTier,
                Priority = rr.Priority
            }).ToList();

            var outputData_T = new List<ProductCategory>();

            var result = prodWithPriority_T.Select(pp => {
                return SummarizeAverageSheetData(pp);
            });
            return result;
        }

        private static ProductCategory GetDefaultRowInSheet(ResolvedReportRowData reportData, int numberOfColumns)
        {
            var defaultRowInSheet = new ProductCategory() {
                PrimaryTierName = reportData.PrimaryTier,
                SecondaryTierName = reportData.SecondaryTier,
                Properties = new List<ExcelCellDto>()
            };

            for (int i = 0; i < numberOfColumns; i++) {
                defaultRowInSheet.Properties.Add(new ExcelCellDto());
            }
            return defaultRowInSheet;
        }

       

        private static ResolvedReportRowData SummarizeAverageSheetData(IGrouping<ResolvedReportRowData, ResolvedReportRowData> pp)
        {
            var objToReturn_R = new ResolvedReportRowData() {
                PrimaryTier = pp.Key.PrimaryTier,
                SecondaryTier = pp.Key.SecondaryTier,
                Priority = pp.Key.Priority,
                timeTaken = new TimeSpan(0),
                Count = 0
            };

            foreach (var item in pp) {
                objToReturn_R.Count++;
                objToReturn_R.timeTaken += item.timeTaken;
            }

            return objToReturn_R;
        }

       

       

      

        internal static List<ResolvedReportRowData> ParseResolvedSheet(ISheet worksheet)
        {
            List<ResolvedReportRowData> resolvedData = new List<ResolvedReportRowData>();

            for (int it_row = 1; it_row < worksheet.LastRowNum; it_row++) {

                // the magic number 10, 11 etc are column number to specifc data in excel sheet
                // while magic number 5 in 'priority > 5' is lowest possible priority
                int priority;
                string tier1 = worksheet.GetRow(it_row ).GetCell( 10).StringCellValue;
                string tier2 = worksheet.GetRow(it_row).GetCell(11).StringCellValue;

                DateTime issueOpen = worksheet.GetRow(it_row ).GetCell(14).DateCellValue;
                DateTime issueClose = worksheet.GetRow(it_row).GetCell(15).DateCellValue;

                if (Int32.TryParse(worksheet.GetRow(it_row).GetCell(6).StringCellValue, out priority) == false) {
                    priority = 0;
                }

                if (priority > 5 || priority < 0)
                    priority = 0;


                resolvedData.Add(new ResolvedReportRowData() {
                    PrimaryTier = tier1,
                    SecondaryTier = tier2,
                    timeTaken = issueClose - issueOpen,
                    Priority = (int)priority

                });

            }
            return resolvedData;
        }

    }
}
