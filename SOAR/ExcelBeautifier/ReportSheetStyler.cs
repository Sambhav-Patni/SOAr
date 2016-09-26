using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class ReportSheetStyler
    {
        public void StyleSheet(ISheet sheet)
        {
            rowsInHeader = 1;
            footerPresent = false;
            switch (sheet.SheetName) {
                case ExcelReportSummaryConfiguration.NameOfScrReportSheet:
                    StyleScrReport(sheet);
                    break;
                case ExcelReportSummaryConfiguration.NameOfAgeFromOpenDateSheet:
                    rowsInHeader = 2;
                    footerPresent = true;
                    StyleAgeFromOpenDateSheet(sheet);
                    break;
                case ExcelReportSummaryConfiguration.NameOfByProductSheet:
                     rowsInHeader = 2;
                     footerPresent = true;
                    StyleByProductSheet(sheet);
                    break;
                case  ExcelReportSummaryConfiguration.NameOfByIncidentStatusSheet:
                    StyleByIncidentStatus(sheet);
                    break;
                case ExcelReportSummaryConfiguration.NameOfLeakageRateSheet:
                    StyleLeakageSheet(sheet);
                    break;

                case ExcelReportSummaryConfiguration.NameOfAverageTimeToResolveSheetInOutput:
                    StyleAverageTimeToResolveSheet(sheet);
                    break;
                default:
                    break;
            }
           
                     


        }
        private void CommonStyles(ISheet sheet)
        {
            headerRowCellFont = (HSSFFont)sheet.Workbook.CreateFont(); // Create a new font in the workbook
            headerRowCellFont.FontName = "Calibri";
            headerRowCellFont.Color = HSSFColor.Black.Index;
            headerRowCellFont.Boldweight = 800;

            HSSFCellStyle headerCellStyle = (HSSFCellStyle)sheet.Workbook.CreateCellStyle();
            AddHeaderStyles(headerCellStyle);
            headerCellStyle.SetFont(headerRowCellFont);

            headerCellStyle.SetFont(headerRowCellFont);

            HSSFCellStyle headerCellStyle1 = (HSSFCellStyle)sheet.Workbook.CreateCellStyle();
            AddHeaderStyles(headerCellStyle1);
            headerCellStyle1.SetFont(headerRowCellFont);

            List<HSSFCellStyle> headerStyles = new List<HSSFCellStyle>();
            headerStyles.Add(headerCellStyle);
            headerStyles.Add(headerCellStyle1);

            //var valueCellStyle = (HSSFCellStyle)sheet.Workbook.CreateCellStyle();
            //AddTotalISsueCellStyle(valueCellStyle);

            //style headers 
            for (int head_it = 0; head_it < rowsInHeader; head_it++) {
                var cells = sheet.GetRow(head_it).Cells;
                foreach (var item in cells) {
                    item.CellStyle = headerCellStyle;
                }
            }
            //add border to excell cells


            //add grey foreground color to first 2 columns
            HSSFCellStyle valueCellStyle = (HSSFCellStyle)sheet.Workbook.CreateCellStyle();
            HSSFCellStyle normalCellStyle = (HSSFCellStyle)sheet.Workbook.CreateCellStyle();
             
            valueCellStyle.FillPattern = FillPattern.SolidForeground;
            valueCellStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            AddThinBorder(valueCellStyle);

            AddThinBorder(normalCellStyle);
            valueCellStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            normalCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;

            // align = false;
            for (int body_it = rowsInHeader; body_it <= sheet.LastRowNum; body_it++) {

                var cell = sheet.GetRow(body_it).GetCell(0);
                if (cell != null) {
                    cell.CellStyle = valueCellStyle;
                }
                cell = sheet.GetRow(body_it).GetCell(1);
                if (cell != null) {
                    cell.CellStyle = valueCellStyle;
                }



                var cells = sheet.GetRow(body_it).Cells.Skip(2).ToList();

                foreach (var item in cells) {
                    item.CellStyle = normalCellStyle;
                }

            }

            //style footer if it exists
            if (footerPresent == true) {
                var footer = sheet.GetRow(sheet.LastRowNum);

                var fcells = footer.Cells.ToList();
                foreach (var item in fcells) {
                    item.CellStyle = headerCellStyle;
                }

            }
            AutoColumnSizeAdjust(sheet);
        }

        private void StyleAverageTimeToResolveSheet(ISheet sheet)
        {
            //add yellow foreground to total issues
           
            CommonStyles(sheet);

            var cellStyle = sheet.GetRow(2).GetCell(3).CellStyle;
            var goldCellstyle = sheet.Workbook.CreateCellStyle();
            var numberStyle = sheet.Workbook.CreateCellStyle();
            var numberFormat =  sheet.Workbook.CreateDataFormat().GetFormat( "0.0" ); 
            goldCellstyle.CloneStyleFrom(cellStyle);
            AddTotalISsueCellStyle(goldCellstyle as HSSFCellStyle);
            
            //last row is footer and hence need not be styled
            for (int row_st = 1; row_st < sheet.LastRowNum; row_st++) {
                for (int col_it = 2; col_it < 8; col_it++) {
                        var cell =   sheet.GetRow(row_st).GetCell(col_it);
                        if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric) {
                            cell.CellStyle.DataFormat = numberFormat;
                        }
                   
                }
                sheet.GetRow(row_st).GetCell(8).CellStyle = goldCellstyle;
               
            }
          //  sheet.AutoSizeColumn(1);
            AutoColumnSizeAdjust(sheet);
        }
        public void AddDocumentInformation(HSSFWorkbook workbook)
        {
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Infogain India Pvt. Ltd.";
            workbook.DocumentSummaryInformation = dsi;
            SummaryInformation docme = PropertySetFactory.CreateSummaryInformation();
            docme.Author = "Raj Kamal";
            docme.Title = "Support Operations Analysis Report";
            docme.Comments = "Created By SOAr\nCredits: Sambhav Patni(sambhav.patni@infogain.com) & Raj Kamal( raj.kamal@infogain.com)";
            workbook.SummaryInformation = docme;
          
        }
        public void AddDocumentInformation(XSSFWorkbook workbook)
        {
            var xmlProps = workbook.GetProperties();
            var coreProps = xmlProps.CoreProperties;
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Infogain India Pvt. Ltd.";
          //  workbook.DocumentSummaryInformation = dsi;
            SummaryInformation docme = PropertySetFactory.CreateSummaryInformation();
            coreProps.Creator = "Raj Kamal";
            coreProps.Title = "Support Operations Analysis Report";
            coreProps.Description = "Created By SOAr\nCredits: Sambhav Patni(sambhav.patni@infogain.com) & Raj Kamal( raj.kamal@infogain.com)";
          //  workbook.SummaryInformation = docme;
          
        }


       

        private static void AutoColumnSizeAdjust(ISheet sheet)
        {
            int columnCount = sheet.GetRow(0).Cells.Count;
            int padding = 256*3;
            for (int i = sheet.LeftCol; i <= columnCount; i++) {
                sheet.AutoSizeColumn(i, true);
                sheet.SetColumnWidth(i, sheet.GetColumnWidth(i) + padding);
            }
        }

        private static void AddTotalISsueCellStyle(HSSFCellStyle totalIssueStyle)
        {
            totalIssueStyle.FillPattern = FillPattern.SolidForeground;
            totalIssueStyle.FillForegroundColor = HSSFColor.LemonChiffon.Index;
           
        }

        private static void AddHeaderStyles(HSSFCellStyle headerCellStyle)
        {
            headerCellStyle.FillForegroundColor = HSSFColor.PaleBlue.Index;//HSSFColor.Aqua.Index;
            headerCellStyle.FillPattern = FillPattern.SolidForeground;
            headerCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
            AddThinBorder(headerCellStyle);
        }

        private static void AddThinBorder(HSSFCellStyle headerCellStyle)
        {
            headerCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            headerCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            headerCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            headerCellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
        }

        private void StyleByProductSheet(ISheet sheet)
        {
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 2, 8));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 9, 16));
            CommonStyles(sheet);

            List<int> goldCells = new List<int>() { 8, 15};

            StyleAsTotalCell(sheet, goldCells);
        }

        private static void StyleAsTotalCell(ISheet sheet, List<int> goldCells)
        {
            var cellStyle = sheet.GetRow(2).GetCell(3).CellStyle;
            var goldCellstyle = sheet.Workbook.CreateCellStyle();
            goldCellstyle.CloneStyleFrom(cellStyle);
            AddTotalISsueCellStyle(goldCellstyle as HSSFCellStyle);
            //last row is footer and hence need not be styled
            for (int row_st = 2; row_st <= sheet.LastRowNum - 1; row_st++) {
                goldCells.ForEach(cell_num => {
                    sheet.GetRow(row_st).GetCell(cell_num).CellStyle = goldCellstyle;
                });
            }
        }


        private void StyleScrReport(ISheet sheet)
        {
            CommonStyles(sheet);
            //AutoColumnSizeAdjust(sheet);
            //nothing extra
        } 

        private void StyleByIncidentStatus(ISheet sheet)
        {
            CommonStyles(sheet);

            //add yellow foreground to total issues

          //  AutoColumnSizeAdjust(sheet);

        }

        private void StyleAgeFromOpenDateSheet(ISheet sheet)
        {

            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 2, 7));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 8, 13));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 14, 19));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 20, 25));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 26, 31));


         
            CommonStyles(sheet);
            //add yellow foreground to total issues
            var goldCells = new List<int>() { 
              32
            };
            StyleAsTotalCell(sheet, goldCells);
        
          //  AutoColumnSizeAdjust(sheet);
           

        }

        private void StyleLeakageSheet(ISheet sheet)
        {
            CommonStyles(sheet);
           
         //   AutoColumnSizeAdjust(sheet);
            

        }
        
        public HSSFFont headerRowCellFont { get; set; }

        public int rowsInHeader { get; set; }

        public bool footerPresent { get; set; }
    }
}
