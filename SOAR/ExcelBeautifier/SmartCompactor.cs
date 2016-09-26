using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class SmartCompactor
    {
        public static void GenerateMonthlyReportSummary(string inputExcelFilePath, string inputExcelResolvedFilePath, string outputExcelFilePath)
        {
            var main = new SmartCompactor();

            FileInfo excelFile = new FileInfo(inputExcelFilePath);

            ExcelPackage excel = new ExcelPackage(excelFile);

            var compact = main.Run(excel);

            ExcelPackage resolvedExcelFile = new ExcelPackage(new FileInfo(inputExcelResolvedFilePath));

            var averageSheet = main.GetAverageResolutionTime(
                                        resolvedExcelFile,
                                        excel.Workbook.Worksheets["Average time to resolve"]
                                );
            compact.Workbook.Worksheets.Add("Average Time To resolve (compact)", averageSheet);

            compact.SaveAs(new FileInfo(outputExcelFilePath));
        }

        public ExcelWorksheet GetAverageResolutionTime(ExcelPackage source, ExcelWorksheet template)
        {
            var worksheet = source.Workbook.Worksheets["resolved"]; // there should be only 1 worksheet in the excel
            List<ResolvedReportRowData> resolvedData = ParseResolvedSheet(worksheet);

            var generalGroups = GetTier1Groups(resolvedData);
            var apdClaimsGroups = GetTier2ApdClaimsGroups(resolvedData);

      
            ExcelPackage sheetContainer = new ExcelPackage();
            var averageSheet = sheetContainer.Workbook.Worksheets.Add("AvgData", template);

            WriteToAverageSheet( averageSheet, generalGroups, apdClaimsGroups);

            int rows = averageSheet.Dimension.End.Row;
            int expectedRows = generalGroups.Count + apdClaimsGroups.Count + 1;
            averageSheet.DeleteRow(expectedRows + 1, rows - expectedRows);

            return averageSheet;
        }

        private static void WriteToAverageSheet( ExcelWorksheet averageSheet, List<IGrouping<ReportData, ResolvedReportRowData>> generalGroups, List<IGrouping<ReportData, ResolvedReportRowData>> apdClaimsGroups)
        {
            int excel_row = 2;
            WriteToWorkSheet(averageSheet, generalGroups, excel_row);
            WriteToWorkSheet(averageSheet, apdClaimsGroups, excel_row + generalGroups.Count);
            averageSheet.Cells[1, 1, averageSheet.Dimension.End.Row, averageSheet.Dimension.End.Column - 1].Style.Numberformat.Format = "0.0";
        }

        private static List<IGrouping<ReportData, ResolvedReportRowData>> GetTier2ApdClaimsGroups(List<ResolvedReportRowData> resolvedData)
        {
            var apdClaimsData = resolvedData.Where(rd => rd.PrimaryTier == ReportConstants.ApdClaimsTier1Name).ToList();
            apdClaimsData.ForEach(acd => {
                acd.SecondaryTier = ReportConstants.GetApdClaimsTier2HeadingParent(acd.SecondaryTier);
            });

            var apdClaimsGroups = apdClaimsData.GroupBy(
                                     acd => new ReportData {
                                         Tier1Name = acd.PrimaryTier,
                                         Tier2Name = acd.SecondaryTier
                                     }
                                 ).ToList();

            return apdClaimsGroups;
        }

        private static List<IGrouping<ReportData, ResolvedReportRowData>> GetTier1Groups(List<ResolvedReportRowData> resolvedData)
        {
            var generalData = resolvedData.Where(rd => rd.PrimaryTier != ReportConstants.ApdClaimsTier1Name).ToList();

            generalData.ForEach(gd => {
                gd.PrimaryTier = ReportConstants.GetTier1HeadingParent(gd.PrimaryTier);
                gd.SecondaryTier = "";

            });



            var generalGroups = generalData.GroupBy(
                                 gd => new ReportData {
                                     Tier1Name = gd.PrimaryTier,
                                     Tier2Name = gd.SecondaryTier
                                 }
                             ).ToList();
            return generalGroups;
        }

        private static List<ResolvedReportRowData> ParseResolvedSheet(ExcelWorksheet worksheet)
        {
            List<ResolvedReportRowData> resolvedData = new List<ResolvedReportRowData>();
            for (int it_row = 2; it_row <= worksheet.Dimension.End.Row; it_row++) {

                // the magic number 11, 12 etc are column number to specifc data in excel sheet
                // while magic number 5 in 'priority > 5' is lowest possible priority
                int priority;
                string tier1 = worksheet.Cells[it_row, 11].Text;
                string tier2 = worksheet.Cells[it_row, 12].Text;
              
                DateTime issueOpen = worksheet.Cells[it_row, 15].GetValue<DateTime>();
                DateTime issueClose = worksheet.Cells[it_row, 16].GetValue<DateTime>();

                if (Int32.TryParse(worksheet.Cells[it_row, 7].Text, out priority) == false) {
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

        private static void WriteToWorkSheet(ExcelWorksheet worksheet, List<IGrouping<ReportData, ResolvedReportRowData>> generalGroups, int excel_row)
        {

            for (int it_ii = 0; it_ii < generalGroups.Count; it_ii++) {

                var priorityWise = generalGroups[it_ii].GroupBy(gg => gg.Priority);
                IGrouping<int, ResolvedReportRowData> data;
                TimeSpan totalTimeTaken;

                int excel_col = 2;
                string tier1 = generalGroups[it_ii].Key.Tier1Name;


                worksheet.Cells[excel_row + it_ii, 1].Value = generalGroups[it_ii].Key.Tier1Name;
                if (it_ii > 0 && generalGroups[it_ii].Key.Tier1Name == generalGroups[it_ii - 1].Key.Tier1Name) {
                    worksheet.Cells[excel_row + it_ii, 1].Value = "";
                }


                worksheet.Cells[excel_row + it_ii, 2].Value = generalGroups[it_ii].Key.Tier2Name;

                for (int it_jj = 1; it_jj <= 5; it_jj++) {

                    AggregateTimeTakenForResolvingIssue(priorityWise, it_jj, out data, out totalTimeTaken);

                    if (data != null && data.Count() != 0)
                        worksheet.Cells[excel_row + it_ii, excel_col + it_jj].Value = totalTimeTaken.TotalDays / data.Count();
                }

                //for out of priority range data               
                AggregateTimeTakenForResolvingIssue(priorityWise, 0, out data, out totalTimeTaken);
                if (data != null && data.Count() != 0)
                    worksheet.Cells[excel_row + it_ii, excel_col + 6].Value = totalTimeTaken.TotalDays / data.Count();

                //total issues
                worksheet.Cells[excel_row + it_ii, excel_col + 7].Value = generalGroups[it_ii].Count();

            }
        }

        private static void AggregateTimeTakenForResolvingIssue(IEnumerable<IGrouping<int, ResolvedReportRowData>> priorityWise, int it_jj, out IGrouping<int, ResolvedReportRowData> data, out TimeSpan totalTimeTaken)
        {

            data = priorityWise.FirstOrDefault(pw => pw.Key == it_jj);

            if (data == null) {
                totalTimeTaken = new TimeSpan(0);
            }
            else {
                totalTimeTaken = data.Select(dt => dt.timeTaken).Aggregate((a, b) => { return a + b; });
            }


        }

        public ExcelPackage Run(ExcelPackage excel)
        {

            // copy excel worksheets to memory.
            var tsheets = excel.Workbook.Worksheets.
                            Where(sheet => ReportConstants.
                                                ReportTabsToModify.Contains(sheet.Name)
                            ).ToList();

            ExcelPackage t_excel = new ExcelPackage();
            ExcelPackage c_excel = new ExcelPackage();

            var sheets = tsheets.Select(sh =>
                                    t_excel.Workbook.Worksheets.Add(sh.Name + " type ", sh)
                        ).ToList();
            
            foreach (var it_Sheet in sheets) {


                var reportData = this.ParseReportsheet(it_Sheet);

                if (reportData.Count == 0)
                    continue;
                int lastrow = reportData.Last().DataRange.Start.Row;

                int row_offset = reportData[0].DataRange.Start.Row; // 2 or 3

                //  List<ReportRow> newReportData = new List<ReportRow>();

                var generalReportData = reportData.
                                            Where(r => r.Tier1Name != ReportConstants.ApdClaimsTier1Name).
                                            ToList();

                var apdClaimsReportData = reportData.
                                            Where(r => r.Tier1Name == ReportConstants.ApdClaimsTier1Name).
                                            ToList();

                if (it_Sheet.Name.Contains("SCR Report")) {
                    ChangeValuesToCount(generalReportData);
                    ChangeValuesToCount(apdClaimsReportData);


                }

                if (it_Sheet.Name.ToLower().Contains("Average time to resolve".ToLower())) {
                    continue;
                }

                var reportMerger = new ReportDataMerger();
                reportMerger.Aggregate = ReportDataMerger.SumExcelCell;

                //merging inner tier entries
                reportMerger.MergerOption = ReportMergeOption.AllIntraTier;
                CompactReportRows(generalReportData, reportMerger);

                //merging miscellanious
                reportMerger.MergerOption = ReportMergeOption.ApdMiscellaniousMerge;
                CompactReportRows(generalReportData, reportMerger);

                //merging claims data
                reportMerger.MergerOption = ReportMergeOption.ApdClaimsIntraTierMerge;
                CompactReportRows(apdClaimsReportData, reportMerger);

                //writing to new sheet

                bool changeHeaders = false;
                if (reportData[0].DataRange.Start.Column != 3) { //data should start at column 3

                    changeHeaders = true;
                }
                var compactSheet = c_excel.Workbook.Worksheets.Add(it_Sheet.Name + " (compact)", it_Sheet);

                if (changeHeaders == true) {
                    UpdateColumnHeaders(reportData, compactSheet);
                }

                //Ensuring that general data has right names
                generalReportData.ForEach(gdr => {
                    gdr.Tier1Name = ReportConstants.GetTier1HeadingParent(gdr.Tier1Name);
                    gdr.Tier2Name = "";
                });

                apdClaimsReportData.ForEach(gdr => {
                    gdr.Tier2Name = ReportConstants.GetApdClaimsTier2HeadingParent(gdr.Tier2Name);

                });


                generalReportData.Sort((a, b) => {

                    int id1 = ReportConstants.Tier1Headings.FindIndex(head => head == a.Tier1Name);
                    int id2 = ReportConstants.Tier1Headings.FindIndex(head => head == b.Tier1Name);
                    return id1.CompareTo(id2);

                });

                apdClaimsReportData.Sort((a, b) => {

                    int id1 = ReportConstants.ApdOpsTier2Headings.FindIndex(head => head == a.Tier2Name);
                    int id2 = ReportConstants.ApdOpsTier2Headings.FindIndex(head => head == b.Tier2Name);
                    return id1.CompareTo(id2);

                });

                FillCompactSheet(generalReportData, compactSheet, row_offset);
                FillCompactSheet(apdClaimsReportData, compactSheet, row_offset + generalReportData.Count);

                int finalDataRow = row_offset + generalReportData.Count + apdClaimsReportData.Count;

                compactSheet.DeleteRow(finalDataRow, lastrow - finalDataRow + 1);

            }

            return c_excel;

        }

        private static void ChangeValuesToCount(List<ReportRow> generalReportData)
        {
            generalReportData.ForEach(rd => {
                for (int it_col = rd.DataRange.Start.Column; it_col <= rd.DataRange.End.Column; it_col++) {
                    rd.DataRange[rd.DataRange.Start.Row, it_col].Value = 1;
                }
            });
        }
        
        private static void UpdateColumnHeaders(List<ReportRow> reportData, ExcelWorksheet compactSheet)
        {
            int columnStart = reportData[0].DataRange.Start.Column - 2;

            //input : 1, 2, 3, 4, 5
            //output : 3, 4, 5

            compactSheet.Cells[1,
                                columnStart,
                                compactSheet.Dimension.End.Row,
                                compactSheet.Dimension.End.Column
            ].Copy(
                compactSheet.Cells[
                    1,
                    1,
                    compactSheet.Dimension.End.Row,
                    compactSheet.Dimension.End.Column - columnStart + 1
                ]
            );
            //deleting extra columns
            if (compactSheet.Dimension.End.Column > columnStart)
                compactSheet.DeleteColumn(compactSheet.Dimension.End.Column - columnStart + 2, columnStart);
        }

        private static void FillCompactSheet(List<ReportRow> generalReportData, ExcelWorksheet compactSheet, int row_offset)
        {

        //    PersistentExcelRange it_sheet = new PersistentExcelRange(compactSheet);
            for (int it_i = 0; it_i < generalReportData.Count; it_i++) {

                if (it_i == 0 ||
                    generalReportData[it_i].Tier1Name != generalReportData[it_i - 1].Tier1Name
                ) {
                    compactSheet.Cells[row_offset + it_i, 1].Value = generalReportData[it_i].Tier1Name;
                }
                else {
                    compactSheet.Cells[row_offset + it_i, 1].Value = "";
                }

                compactSheet.Cells[row_offset + it_i, 2].Value = generalReportData[it_i].Tier2Name;

                var columnData = generalReportData[it_i].DataRange;
                int col_offset = 3;
                for (int it_col = 0; it_col <= columnData.End.Column - columnData.Start.Column; it_col++) {

                    compactSheet.
                        Cells[row_offset + it_i, col_offset + it_col].
                        Value = columnData[it_col, true].Value;
                }

            }
        }
       
        private static void CompactReportRows(List<ReportRow> listOfUncompressedReportRows, ReportDataMerger reportMerger)
        {

          //effectively we are merging entries which can be emerged to others
          //using simpler technique 

            List<ReportRow> listOfCompressedReports = new List<ReportRow>();         
        
            foreach (var uncompressedReport in listOfUncompressedReportRows) {
                bool merged = false;

                foreach (var report in listOfCompressedReports) {
                    //SumReport() merges if possible.
                    if (reportMerger.SumReportRows(report, uncompressedReport) == true) {
                       merged = true;
                       break;
                    }
                }

                if (merged == false) {
                    listOfCompressedReports.Add(uncompressedReport);
                }
            }

            listOfUncompressedReportRows.Clear();
            listOfUncompressedReportRows.AddRange(listOfCompressedReports);

        }



        public List<ReportRow> ParseReportsheet(ExcelWorksheet source)
        {
            //find tier1 and tier2 columns

            bool dataExists = false;
            int startRow = 0, startCol = 0;
            for (int it_row = 1; it_row < source.Dimension.End.Row; it_row++) {
                for (int it_col = 1; it_col < source.Dimension.End.Column; it_col++) {
                    if (source.Cells[it_row, it_col].Text == ReportConstants.Tier1ColumnHeading &&
                         source.Cells[it_row, it_col + 1].Text == ReportConstants.Tier2ColumnHeading
                    ) {
                        startRow = it_row + 1; //data will definitely starts from next row instead of current row

                        if (source.Cells[startRow, it_col].Text == "") //if next cell is blank instead of tier1 name
                            {
                            startRow++;
                        }
                        startCol = it_col;
                        dataExists = true;
                        break;
                    }
                }
                if (dataExists == true)
                    break;
            }

            if (dataExists == false) {
                throw new DataNotFoundException();
            }

          
            var targetRange = new PersistentExcelRange();
            targetRange.setRange(source, startRow, startCol, source.Dimension.End.Row, source.Dimension.End.Column);
          
            //create and fill reportRows
            List<ReportRow> reportRows = new List<ReportRow>();
            for (int it_row = 0; it_row <= targetRange.Range.End.Row - targetRange.Range.Start.Row; it_row++) {
                //tier 1 name is present in column 1(offset 0), tier 2 name in column 2(offset 1)     

                string tier1name = targetRange[it_row, 0, true].GetValue<string>();
                string tier2name = targetRange[it_row, 1, true].GetValue<string>();

                if (tier1name == null && tier2name == "Total" && it_row == targetRange.Range.End.Row - targetRange.Range.Start.Row) {
                    //this is the total line (which contains Sum of all column data) and should not be considered
                    continue;
                }

                if (tier1name == "" || tier1name == null) {
                    tier1name = reportRows.Last().Tier1Name;
                }
                if (tier2name == "" || tier2name == null) {
                    tier2name = reportRows.Last().Tier2Name;
                }

                int lastColumnToUse = targetRange.Range.End.Column;

                if (source.Cells[targetRange.Range.Start.Row + it_row, targetRange.Range.End.Column].IsFormula() == true) {
                    lastColumnToUse--;
                }               

                reportRows.Add(new ReportRow() {
                    Tier1Name = tier1name,
                    Tier2Name = tier2name,
                    DataRange = new PersistentExcelRange(source,
                                      targetRange.Range.Start.Row + it_row,
                                      targetRange.Range.Start.Column + 2, // data starts at +2 offset
                                      targetRange.Range.Start.Row + it_row,
                                      lastColumnToUse
                    )
                });
            }

            return reportRows;

        }        
    }
}
