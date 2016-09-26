using ExcelLibrary.SpreadSheet;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using ExcelBeautifier;

namespace SOAR_Sam
{
    public partial class Main : Form
    {
        string path;        
        string[] Month = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
        public Main()
        {             
            InitializeComponent();
            int year = DateTime.Now.Year;
            if (DateTime.Now.Month == 1)
            {
                year--;
            }
            dateTimePicker1.Value = dateTimePicker1.Value.Subtract(TimeSpan.FromDays(30));
            textBox4.Text = Month[DateTime.Now.Month - 1];
            //label9.Text = Month[DateTime.Now.Month - 2] + year + "_APD-OPS";
            textBox5.Text = Month[DateTime.Now.Month - 1] + "_" + year + "_APD_OPS";
        }        

        public void ExportDataTableToExcel(string fileName, DataSet sourceSet, bool color, string month)
        {
            //label18.Show();
            label18.Invoke((MethodInvoker)(() => label18.Show()));
            FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
            HSSFWorkbook workbook = new HSSFWorkbook();                        
            workbook.CreateSheet("Backlog");

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Infogain India Pvt. Ltd.";
            workbook.DocumentSummaryInformation = dsi;
            SummaryInformation docme = PropertySetFactory.CreateSummaryInformation();
            docme.Author = "Sambhav Patni";
            docme.Title = "Support Operations Analysis Report";
            docme.Comments = "Created By SOAr\nCredits: Sambhav Patni(sambhav.patni@infogain.com)";
            //workbook.SummaryInformation.Author = "Sambhav Patni";
            //workbook.SummaryInformation.Title = "Support Operations Analysis Report";
            //workbook.SummaryInformation.Comments = "Created By SOAr\nCredits: Sambhav Patni(sambhav.patni@infogain.com)";
            workbook.SummaryInformation = docme;
            foreach (DataTable sourceTable in sourceSet.Tables)
            {
                HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet(sourceTable.TableName);
                int rowIndex = 2;
                if (sourceTable.TableName.Equals("By Product") || sourceTable.TableName.Equals("By Customer") || sourceTable.TableName.Equals("By Region") || sourceTable.TableName.Equals("Age From Open Date") || sourceTable.TableName.Equals("By Incident Status"))
                {
                    rowIndex = 2;
                }
                else
                {
                    rowIndex = 1;
                }
                HSSFFont headerRowCellFont = (HSSFFont)workbook.CreateFont(); // Create a new font in the workbook
                headerRowCellFont.FontName = "Calibri";
                headerRowCellFont.Color = HSSFColor.Black.Index;
                headerRowCellFont.Boldweight = 800;
                // handling header.
                if (sourceTable.TableName == "Age From Open Date")
                {
                    HSSFCellStyle headerCellStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                    headerCellStyle.FillForegroundColor = HSSFColor.PaleBlue.Index;//HSSFColor.Aqua.Index;
                    headerCellStyle.FillPattern = FillPattern.SolidForeground;
                    headerCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
                    headerCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    headerCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    headerCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Hair;
                    headerCellStyle.SetFont(headerRowCellFont);

                    HSSFCellStyle headerCellStyle1 = (HSSFCellStyle)workbook.CreateCellStyle();
                    headerCellStyle1.FillForegroundColor = HSSFColor.PaleBlue.Index;//HSSFColor.Aqua.Index;
                    headerCellStyle1.FillPattern = FillPattern.SolidForeground;
                    headerCellStyle1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
                    headerCellStyle1.SetFont(headerRowCellFont);

                    HSSFRow headerRow1 = (HSSFRow)sheet.CreateRow(0);
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 2, 7));
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 8, 13));
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 14, 19));
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 20, 25));
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 26, 31));
                    HSSFCell headerCell0 = (HSSFCell)headerRow1.CreateCell(0);
                    HSSFCell headerCell1 = (HSSFCell)headerRow1.CreateCell(1);
                    HSSFCell headerCell2 = (HSSFCell)headerRow1.CreateCell(2);
                    HSSFCell headerCell3 = (HSSFCell)headerRow1.CreateCell(8);
                    HSSFCell headerCell4 = (HSSFCell)headerRow1.CreateCell(14);
                    HSSFCell headerCell5 = (HSSFCell)headerRow1.CreateCell(20);
                    HSSFCell headerCell6 = (HSSFCell)headerRow1.CreateCell(26);
                    HSSFCell headerCell7 = (HSSFCell)headerRow1.CreateCell(32);
                    headerCell0.SetCellValue("Product Categorization Tier1");
                    headerCell1.SetCellValue("Product Categorization Tier2");
                    headerCell2.SetCellValue("Due Days > 30");
                    headerCell3.SetCellValue("Due Days > 15");
                    headerCell4.SetCellValue("Due Days > 7");
                    headerCell5.SetCellValue("Due Days > 3");
                    headerCell6.SetCellValue("Due Days > 1");
                    headerCell7.SetCellValue("Total Issues");
                    headerCell0.CellStyle = headerCellStyle;
                    headerCell1.CellStyle = headerCellStyle;
                    headerCell2.CellStyle = headerCellStyle;
                    headerCell3.CellStyle = headerCellStyle;
                    headerCell4.CellStyle = headerCellStyle;
                    headerCell5.CellStyle = headerCellStyle;
                    headerCell6.CellStyle = headerCellStyle;
                    headerCell7.CellStyle = headerCellStyle;
                    sheet.AutoSizeColumn(1);
                    sheet.AutoSizeColumn(2);

                    HSSFRow headerRow2 = (HSSFRow)sheet.CreateRow(1);
                    foreach (DataColumn column in sourceTable.Columns)
                    {
                        HSSFCell headerCell21 = (HSSFCell)headerRow2.CreateCell(column.Ordinal);
                        if (column.Ordinal == 2 || column.Ordinal == 8 || column.Ordinal == 14 || column.Ordinal == 20 || column.Ordinal == 26)
                        {
                            headerCell21.SetCellValue("P1");
                        }
                        if (column.Ordinal == 3 || column.Ordinal == 9 || column.Ordinal == 15 || column.Ordinal == 21 || column.Ordinal == 27)
                        {
                            headerCell21.SetCellValue("P2");
                        }
                        if (column.Ordinal == 4 || column.Ordinal == 10 || column.Ordinal == 16 || column.Ordinal == 22 || column.Ordinal == 28)
                        {
                            headerCell21.SetCellValue("P3");
                        }
                        if (column.Ordinal == 5 || column.Ordinal == 11 || column.Ordinal == 17 || column.Ordinal == 23 || column.Ordinal == 29)
                        {
                            headerCell21.SetCellValue("P4");
                        }
                        if (column.Ordinal == 6 || column.Ordinal == 12 || column.Ordinal == 18 || column.Ordinal == 24 || column.Ordinal == 30)
                        {
                            headerCell21.SetCellValue("P5");
                        }
                        if (column.Ordinal == 7 || column.Ordinal == 13 || column.Ordinal == 19 || column.Ordinal == 25 || column.Ordinal == 31)
                        {
                            headerCell21.SetCellValue("PB");
                        }
                        if(column.Ordinal==32)
                        {
                            headerCell21.SetCellValue(" ");
                        }
                        headerCell21.CellStyle = headerCellStyle1;
                        headerCell21.CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
                    }
                }
                else if (sourceTable.TableName == "By Product" || sourceTable.TableName == "By Customer" || sourceTable.TableName == "By Region" || sourceTable.TableName == "By Incident Status")
                {
                    HSSFCellStyle headerCellStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                    headerCellStyle.FillForegroundColor = HSSFColor.PaleBlue.Index;//HSSFColor.Aqua.Index;
                    headerCellStyle.FillPattern = FillPattern.SolidForeground;
                    headerCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
                    headerCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    headerCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    headerCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Hair;
                    headerCellStyle.SetFont(headerRowCellFont);

                    HSSFCellStyle headerCellStyle1 = (HSSFCellStyle)workbook.CreateCellStyle();
                    headerCellStyle1.FillForegroundColor = HSSFColor.PaleBlue.Index;//HSSFColor.Aqua.Index;
                    headerCellStyle1.FillPattern = FillPattern.SolidForeground;
                    headerCellStyle1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
                    headerCellStyle1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
                    headerCellStyle1.SetFont(headerRowCellFont);

                    HSSFRow headerRow1 = (HSSFRow)sheet.CreateRow(0);
                    HSSFCell headerCell0;
                    HSSFCell headerCell1;
                    HSSFCell headerCell2;
                    HSSFCell headerCell3;
                    if (sourceTable.TableName == "By Customer" || sourceTable.TableName == "By Region")
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 1, 7));
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 8, 15));
                        headerCell0 = (HSSFCell)headerRow1.CreateCell(0);
                        headerCell1 = (HSSFCell)headerRow1.CreateCell(1);
                        headerCell2 = (HSSFCell)headerRow1.CreateCell(8);
                        headerCell3 = (HSSFCell)headerRow1.CreateCell(9);
                        headerCell1.SetCellValue("Assigned Incidents(" + month + ")");
                        headerCell2.SetCellValue("Resolved Incidents(" + month + ")");
                    }
                    else if (sourceTable.TableName == "By Incident Status")
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 2, 8));
                        headerCell0 = (HSSFCell)headerRow1.CreateCell(0);
                        headerCell1 = (HSSFCell)headerRow1.CreateCell(1);
                        headerCell2 = (HSSFCell)headerRow1.CreateCell(2);
                        headerCell3 = (HSSFCell)headerRow1.CreateCell(9);
                        headerCell2.SetCellValue("Incident Status for Opened Claims(" + month + ")");
                    }
                    else
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 2, 8));
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 9, 16));
                        headerCell0 = (HSSFCell)headerRow1.CreateCell(0);
                        headerCell1 = (HSSFCell)headerRow1.CreateCell(1);
                        headerCell2 = (HSSFCell)headerRow1.CreateCell(2);
                        headerCell3 = (HSSFCell)headerRow1.CreateCell(9);
                        headerCell2.SetCellValue("Assigned Incidents(" + month + ")");
                        headerCell3.SetCellValue("Resolved Incidents(" + month + ")");
                    }
                    headerCell0.CellStyle = headerCellStyle;
                    headerCell1.CellStyle = headerCellStyle;
                    headerCell2.CellStyle = headerCellStyle;
                    headerCell3.CellStyle = headerCellStyle;

                    HSSFRow headerRow2 = (HSSFRow)sheet.CreateRow(1);
                    foreach (DataColumn column in sourceTable.Columns)
                    {
                        HSSFCell headerCell = (HSSFCell)headerRow2.CreateCell(column.Ordinal);
                        headerCell.SetCellValue(column.ColumnName);
                        headerCell.CellStyle = headerCellStyle1;
                        sheet.AutoSizeColumn(column.Ordinal);
                    }
                }
                else
                {
                    HSSFRow headerRow = (HSSFRow)sheet.CreateRow(0);
                    foreach (DataColumn column in sourceTable.Columns)
                    {
                        //headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                        HSSFCell headerCell = (HSSFCell)headerRow.CreateCell(column.Ordinal);
                        // HSSFRichTextString HeadValue = new HSSFRichTextString(column.ColumnName);                
                        //HeadValue.ApplyFont(2);
                        // Set Cell Value
                        headerCell.SetCellValue(column.ColumnName);
                        HSSFCellStyle headerCellStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                        headerCellStyle.FillForegroundColor = HSSFColor.PaleBlue.Index;//HSSFColor.Aqua.Index;
                        headerCellStyle.FillPattern = FillPattern.SolidForeground;
                        headerCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
                        headerCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
                        headerCellStyle.WrapText = true;
                        headerCell.CellStyle = headerCellStyle;
                        headerCellStyle.SetFont(headerRowCellFont);
                        sheet.AutoSizeColumn(column.Ordinal);
                    }
                }
                //result row
                if (sourceTable.TableName.Equals("By Product") || sourceTable.TableName.Equals("By Customer") || sourceTable.TableName.Equals("By Region") || sourceTable.TableName.Equals("Environment Metrics") || sourceTable.TableName.Equals("Age From Open Date") || sourceTable.TableName.Equals("By Incident Status"))
                {
                    HSSFCellStyle headerCellStyle1 = (HSSFCellStyle)workbook.CreateCellStyle();
                    headerCellStyle1.FillForegroundColor = HSSFColor.PaleBlue.Index;//HSSFColor.Aqua.Index;
                    headerCellStyle1.FillPattern = FillPattern.SolidForeground;
                    headerCellStyle1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
                    headerCellStyle1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
                    headerCellStyle1.SetFont(headerRowCellFont);

                    HSSFRow dataRowTotal = (HSSFRow)sheet.CreateRow(sourceTable.Rows.Count + 2);
                    foreach (DataColumn column in sourceTable.Columns)
                    {
                        HSSFCell valueCell = (HSSFCell)dataRowTotal.CreateCell(column.Ordinal);
                        if (column.ColumnName.Equals("Product Categorization Tier2") || column.ColumnName.Equals("Carrier") || column.ColumnName.Equals("Region/Area") || column.ColumnName.Equals("Environment"))
                        {
                            valueCell.SetCellValue("Total");
                        }
                        else
                        {
                            if (column.ColumnName.Substring(0, 2).Equals("P1") || column.ColumnName.Substring(0, 2).Equals("P2") || column.ColumnName.Substring(0, 2).Equals("P3") || column.ColumnName.Substring(0, 2).Equals("P4") || column.ColumnName.Substring(0, 2).Equals("P5") || column.ColumnName.Substring(0, 2).Equals("PB") || column.ColumnName.Equals("Blank") || column.ColumnName.Equals("BlankR") || column.ColumnName.Equals("Total Issues") || column.ColumnName.Equals("Issues Resolved") || column.ColumnName.Equals("ASSIGNED") || column.ColumnName.Equals("IN PROGRESS") || column.ColumnName.Equals("OPENED") || column.ColumnName.Equals("PENDING") || column.ColumnName.Equals("WAITING FOR RESPONSE") || column.ColumnName.Equals("RESOLVED"))
                                // || column.ColumnName.Equals("P1R") || column.ColumnName.Equals("P2R") || column.ColumnName.Equals("P3R") || column.ColumnName.Equals("P4R") || column.ColumnName.Equals("P5R"))
                            {
                                string[] colpos = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH" };
                                valueCell.SetCellFormula("SUM(" + colpos[column.Ordinal] + (rowIndex + 1) + ":" + colpos[column.Ordinal] + dataRowTotal.RowNum + ")");
                            }

                            if (column.ColumnName.Equals("% Resolved"))
                            {
                                if (column.Table.TableName.Equals("By Product"))
                                {
                                    valueCell.SetCellFormula("CEILING((P" + (dataRowTotal.RowNum + 1) + "*100/I" + (dataRowTotal.RowNum + 1) + "),1)");
                                }
                                if (column.Table.TableName.Equals("By Customer"))
                                {
                                    valueCell.SetCellFormula("CEILING((O" + (dataRowTotal.RowNum + 1) + "*100/H" + (dataRowTotal.RowNum + 1) + "),1)");
                                }
                                if (column.Table.TableName.Equals("By Region"))
                                {
                                    valueCell.SetCellFormula("CEILING((O" + (dataRowTotal.RowNum + 1) + "*100/H" + (dataRowTotal.RowNum + 1) + "),1)");
                                }
                            }
                        }
                        valueCell.CellStyle = headerCellStyle1;
                        sheet.AutoSizeColumn(column.Ordinal);
                    }
                }
                // handling value.
                foreach (DataRow row in sourceTable.Rows)
                {
                    HSSFRow dataRow = (HSSFRow)sheet.CreateRow(rowIndex);
                    bool align = true;
                    foreach (DataColumn column in sourceTable.Columns)
                    {
                        HSSFCellStyle valueCellStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                        HSSFCell valueCell = (HSSFCell)dataRow.CreateCell(column.Ordinal);
                        if (column.ColumnName.Equals("% Resolved"))
                        {
                            //valueCell.SetCellType(CellType.Formula);
                            if (column.Table.TableName.Equals("By Product"))
                            {
                                valueCell.SetCellFormula("CEILING((P" + (rowIndex + 1) + "*100/I" + (rowIndex + 1) + "),1)");
                            }
                            if (column.Table.TableName.Equals("By Customer"))
                            {
                                valueCell.SetCellFormula("CEILING((O" + (rowIndex + 1) + "*100/H" + (rowIndex + 1) + "),1)");
                            }
                            if (column.Table.TableName.Equals("By Region"))
                            {
                                valueCell.SetCellFormula("CEILING((O" + (rowIndex + 1) + "*100/H" + (rowIndex + 1) + "),1)");
                            }
                        }
                        else if (column.ColumnName.Equals("Product Categorization Tier1"))
                        {
                            if (sourceTable.TableName.Equals("By Product") || sourceTable.TableName.Equals("By Customer") || sourceTable.TableName.Equals("By Region") || sourceTable.TableName.Equals("Age From Open Date") || sourceTable.TableName.Equals("By Incident Status"))
                            {
                                try
                                {
                                    if ((row[column]).Equals(DBNull.Value))
                                    {
                                        valueCell.SetCellValue("Not Specified");
                                    }                                   
                                    else if (((string)sourceTable.Rows[dataRow.RowNum - 3]["Product Categorization Tier1"]).Equals(row[column].ToString()))
                                    {
                                        valueCell.SetCellValue(string.Empty);
                                    }
                                    else
                                    {                                       
                                        valueCell.SetCellValue((string)row[column]);                                       
                                    }
                                }
                                catch (InvalidCastException)
                                {
                                    valueCell.SetCellValue((string)row[column]);
                                }
                                catch (IndexOutOfRangeException)
                                {
                                    valueCell.SetCellValue(row[column].ToString());
                                }
                            }
                            else
                            {
                                try
                                {
                                    if ((row[column]).Equals(DBNull.Value))
                                    {
                                        valueCell.SetCellValue("Not Specified");
                                    }  
                                    else if (((string)sourceTable.Rows[dataRow.RowNum - 2]["Product Categorization Tier1"]).Equals(row[column].ToString()))
                                    {
                                        valueCell.SetCellValue(string.Empty);
                                    }
                                    else
                                    {
                                        valueCell.SetCellValue((string)row[column]);                                       
                                    }
                                }
                                catch (InvalidCastException)
                                {
                                    valueCell.SetCellValue((string)row[column]);
                                }
                                catch (IndexOutOfRangeException)
                                {
                                    valueCell.SetCellValue(row[column].ToString());
                                }
                            }
                        }
                        else if (column.ColumnName.Equals("Product Categorization Tier2"))
                        {
                            if (sourceTable.TableName.Equals("By Product") || sourceTable.TableName.Equals("By Customer") || sourceTable.TableName.Equals("By Region") || sourceTable.TableName.Equals("Age From Open Date") || sourceTable.TableName.Equals("By Incident Status"))
                            {
                                try
                                {
                                    if ((row[column]).Equals(DBNull.Value))
                                    {
                                        valueCell.SetCellValue("Not Specified");
                                    }  
                                    else if (((string)sourceTable.Rows[dataRow.RowNum - 3]["Product Categorization Tier2"]).Equals(row[column].ToString()))
                                    {
                                        valueCell.SetCellValue(string.Empty);
                                    }
                                    else
                                    {                                       
                                        valueCell.SetCellValue((string)row[column]);                                       
                                    }
                                }
                                catch (InvalidCastException)
                                {
                                    valueCell.SetCellValue(row[column].ToString());
                                }
                                catch (IndexOutOfRangeException)
                                {
                                    valueCell.SetCellValue(row[column].ToString());
                                }
                            }
                            else
                            {
                                try
                                {
                                    if ((row[column]).Equals(DBNull.Value))
                                    {
                                        valueCell.SetCellValue("Not Specified");
                                    } 
                                    else if (((string)sourceTable.Rows[dataRow.RowNum - 2]["Product Categorization Tier2"]).Equals(row[column].ToString()))
                                    {
                                        valueCell.SetCellValue(string.Empty);
                                    }
                                    else
                                    {
                                        valueCell.SetCellValue((string)row[column]);                                        
                                    }
                                }
                                catch (InvalidCastException)
                                {
                                    valueCell.SetCellValue(row[column].ToString());
                                }
                                catch (IndexOutOfRangeException)
                                {
                                    valueCell.SetCellValue(row[column].ToString());
                                }
                            }
                        }
                        else if (column.ColumnName.Equals("Environment") || column.ColumnName.Equals("Carrier") || column.ColumnName.Equals("Region/Area"))
                        {
                            if ((row[column]).Equals(DBNull.Value))
                            {
                                valueCell.SetCellValue("Not Specified");
                            }
                            else
                            {
                                valueCell.SetCellValue((string)row[column]);
                            }
                        }
                        else
                        {
                            try
                            {
                                valueCell.SetCellValue(Convert.ToDouble(row[column].ToString()));
                                if (valueCell.NumericCellValue.Equals(0))
                                {
                                    valueCell.SetCellValue(string.Empty);           //Cell set 0 cell value Empty
                                }
                            }
                            catch (FormatException)
                            {
                                valueCell.SetCellValue(row[column].ToString());
                            }
                        }
                        //colouring columns
                        try
                        {
                            if (column.ColumnName.Substring(0, 2).Equals("P1") && color == true)
                            {
                                valueCellStyle.FillPattern = FillPattern.SolidForeground;
                                valueCellStyle.FillForegroundColor = HSSFColor.Gold.Index;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                            }
                            if (column.ColumnName.Substring(0, 2).Equals("P2") && color == true)
                            {
                                valueCellStyle.FillPattern = FillPattern.SolidForeground;
                                valueCellStyle.FillForegroundColor = HSSFColor.Orange.Index;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                            }
                            if (column.ColumnName.Substring(0, 2).Equals("P3") && color == true)
                            {
                                valueCellStyle.FillPattern = FillPattern.SolidForeground;
                                valueCellStyle.FillForegroundColor = HSSFColor.Coral.Index;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                            }
                            if (column.ColumnName.Substring(0, 2).Equals("P4") && color == true)
                            {
                                valueCellStyle.FillPattern = FillPattern.SolidForeground;
                                valueCellStyle.FillForegroundColor = HSSFColor.LightBlue.Index;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                            }
                            if (column.ColumnName.Substring(0, 2).Equals("P5") && color == true)
                            {
                                valueCellStyle.FillPattern = FillPattern.SolidForeground;
                                valueCellStyle.FillForegroundColor = HSSFColor.LightGreen.Index;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                            }
                            if ((column.ColumnName.Substring(0, 2).Equals("PB") || column.ColumnName.Substring(0, 2).Equals("Bl")) && color == true)
                            {
                                valueCellStyle.FillPattern = FillPattern.SolidForeground;
                                valueCellStyle.FillForegroundColor = HSSFColor.COLOR_NORMAL;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                            }
                            if (column.ColumnName.Equals("Total Issues") || column.ColumnName.Equals("Issues Resolved"))
                            {
                                valueCellStyle.FillPattern = FillPattern.SolidForeground;
                                valueCellStyle.FillForegroundColor = HSSFColor.LemonChiffon.Index;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                            }
                            if (column.ColumnName.Equals("Product Categorization Tier1") || column.ColumnName.Equals("Product Categorization Tier2") || column.ColumnName.Equals("Region/Area") || column.ColumnName.Equals("Carrier"))
                            {
                                valueCellStyle.FillPattern = FillPattern.SolidForeground;
                                valueCellStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                                align = false;
                            }
                            if (column.ColumnName.Equals("SCR Details"))
                            {
                                align = false;
                                valueCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
                                valueCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
                            }
                        }
                        catch (ArgumentException) { }
                        valueCell.CellStyle = valueCellStyle;
                        if (align != false)
                        {
                            valueCell.CellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CenterSelection;
                        }
                        else
                        {
                            //valueCell.CellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.General;
                            align = true;
                        }
                        sheet.AutoSizeColumn(column.Ordinal);
                    }
                    rowIndex++;
                }
            }
            workbook.CreateSheet("SCR Trend");
            workbook.Write(fs);
            fs.Close();
            //label18.Hide();
            label18.Invoke((MethodInvoker)(() => label18.Hide()));
        }
                
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            textBox1.Text = openFileDialog1.FileName;            
        }
        private void button2_Click(object sender, EventArgs e) 
        {            
            backgroundWorker1.RunWorkerAsync();
        }        
        /*
        {
            label9.Text = "";
            label10.Text = "";
            label11.Text = "";
            label12.Text = "";
            label13.Text = "";
            label14.Text = "";
            label15.Text = "";
            label16.Text = "";
            label17.Text = "";
            if (textBox5.Text == "")
            {
                MessageBox.Show("Please Fill Up The File Name.", "ain't u missin file name Bro...");
                goto Exit_Report;
            }
            string filename = "\\" + textBox5.Text + ".xls";
            label4.Text = "Main: Assigned incidents in the specified duration";
            label5.Text = "Resolved incidents in the specified duration irrespective of when opened";
            label6.Text = "Presently open incidents (not resolved/closed/cancelled) irrespective of when opened";
            label4.ForeColor = Color.Green;
            label5.ForeColor = Color.Green;
            label6.ForeColor = Color.Green;
            try
            {
                if (checkBox8.Checked)
                {
                    if (File.Exists(textBox1.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox1.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                    if (File.Exists(textBox2.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox2.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                    if (File.Exists(textBox3.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox3.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                }
                DataSet DsF = new DataSet();

                if (checkBox1.Checked == true)
                {
                    DsF.Tables.Add(ByProduct().ToTable());
                }
                if (checkBox2.Checked == true)
                {
                    DsF.Tables.Add(ByCustomer().ToTable());
                }
                if (checkBox3.Checked == true)
                {
                    DsF.Tables.Add(SCR().ToTable());
                }
                if (checkBox4.Checked == true)
                {
                    DsF.Tables.Add(ByArea().ToTable());
                }
                if (checkBox5.Checked == true)
                {
                    DsF.Tables.Add(AvgResolveTime().ToTable());
                }
                if (checkBox6.Checked == true)
                {
                    DsF.Tables.Add(ByAgeing().ToTable());
                }
                if (checkBox9.Checked == true)
                {
                    DsF.Tables.Add(EnvMetrics().ToTable());
                }
                if (checkBox10.Checked == true)
                {
                    DsF.Tables.Add(IncidentStatus().ToTable());
                }
                if (checkBox11.Checked == true)
                {
                    DsF.Tables.Add(LeakageRate().ToTable());
                }
                int attempt = 0;
            fl:
                if (File.Exists(Path.GetDirectoryName(path) + filename))
                {
                    filename = filename.Replace("(" + attempt + ")", "");
                    attempt++;
                    filename = filename.Replace(".", "(" + attempt + ").");
                    goto fl;
                    //diaRes = MessageBox.Show("File Already Exists...", "No Big Deal...", MessageBoxButtons.AbortRetryIgnore);
                    //if (diaRes.Equals(DialogResult.Ignore)) { }
                    //if (diaRes.Equals(DialogResult.Abort)) { throw new IOException(); }
                    //if (diaRes.Equals(DialogResult.Retry)) { goto fl; }
                }
                //CreateWorkbook(Path.GetDirectoryName(textBox1.Text) + "\\SOAReport.xls", DsF);                
                ExportDataTableToExcel(Path.GetDirectoryName(path) + filename, DsF, checkBox7.Checked, textBox4.Text);
                if (MessageBox.Show("Report Generated : \n" + Path.GetDirectoryName(path) + filename + "\n\nOpen it??", "Hey Look I have Done Everything", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                {
                    Process.Start(Path.GetDirectoryName(path) + filename);
                }
            }
            catch (IndexOutOfRangeException ex)
            {
                if (!File.Exists(textBox1.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    label4.Text = "Main : File Missing";
                    label4.ForeColor = Color.Red;
                }
                else if (!File.Exists(textBox2.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    label5.Text = "Resolved Sheet : File Missing";
                    label5.ForeColor = Color.Red;
                }
                else if (!File.Exists(textBox3.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    label6.Text = "Age From Open Date : File Missing";
                    label6.ForeColor = Color.Red;
                }
                else
                {
                    MessageBox.Show("Something Went Wrong...\n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com\n\n" + ex.Message, "Error Occured");
                }

            }
            catch (FileNotFoundException)
            {
                if (!File.Exists(textBox1.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    label4.Text = "Main : File Missing";
                    label4.ForeColor = Color.Red;
                }
                else if (!File.Exists(textBox2.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    label5.Text = "Resolved Sheet : File Missing";
                    label5.ForeColor = Color.Red;
                }
                else if (!File.Exists(textBox3.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    label6.Text = "Age From Open Date : File Missing";
                    label6.ForeColor = Color.Red;
                }
            }
            catch (IOException)
            {
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("Excel Not in Proper Format:\n" + ex.Message + ", \n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com", "Error Occured");
            }
            catch (Exception ex)
            {
                if (!File.Exists(textBox1.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    label4.Text = "Main : File Missing";
                    label4.ForeColor = Color.Red;
                }
                else if (!File.Exists(textBox2.Text) && ex.StackTrace.Contains("AvgResolveTime()"))
                {
                    label5.Text = "Resolved Sheet : File Missing";
                    label5.ForeColor = Color.Red;
                }
                else if (!File.Exists(textBox3.Text) && ex.StackTrace.Contains("ByAgeing()"))
                {
                    label6.Text = "Age From Open Date : File Missing";
                    label6.ForeColor = Color.Red;
                }
                else
                {
                    MessageBox.Show("Something Went Wrong...\n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com\n\n" + ex.Message, "Error Occured");
                }
            }
        Exit_Report: ;
        }*/
        private void Do_Work(object sender, DoWorkEventArgs e)
        {
            button2.Invoke((MethodInvoker)(() => button2.Enabled = false));
            //label9.Text = "";
            //label10.Text = "";
            //label11.Text = "";
            //label12.Text = "";
            //label13.Text = "";
            //label14.Text = "";
            //label15.Text = "";
            //label16.Text = "";
            //label17.Text = "";
            label9.Invoke((MethodInvoker)(() => label9.Text = ""));
            label10.Invoke((MethodInvoker)(() => label10.Text = ""));
            label11.Invoke((MethodInvoker)(() => label11.Text = ""));
            label12.Invoke((MethodInvoker)(() => label12.Text = ""));
            label13.Invoke((MethodInvoker)(() => label13.Text = ""));
            label14.Invoke((MethodInvoker)(() => label14.Text = ""));
            label15.Invoke((MethodInvoker)(() => label15.Text = ""));
            label16.Invoke((MethodInvoker)(() => label16.Text = ""));
            label17.Invoke((MethodInvoker)(() => label17.Text = ""));
            label22.Invoke((MethodInvoker)(() => label22.Text = ""));
            if (textBox5.Text == "")
            {
                MessageBox.Show("Please Fill Up The File Name.", "ain't u missin file name Bro...");
                goto Exit_Report;
            }
            string filename = "\\" + textBox5.Text + ".xls";
            //label4.Text = "Main: Assigned incidents in the specified duration";
            //label5.Text = "Resolved incidents in the specified duration irrespective of when opened";
            //label6.Text = "Presently open incidents (not resolved/closed/cancelled) irrespective of when opened";
            //label4.ForeColor = Color.Green;
            //label5.ForeColor = Color.Green;
            //label6.ForeColor = Color.Green;
            label4.Invoke((MethodInvoker)(() => label4.Text = "Main: Assigned incidents in the specified duration"));
            label5.Invoke((MethodInvoker)(() => label5.Text = "Resolved incidents in the specified duration irrespective of when opened"));
            label6.Invoke((MethodInvoker)(() => label6.Text = "Presently open incidents (not resolved/closed/cancelled) irrespective of when opened"));
            label4.Invoke((MethodInvoker)(() => label4.ForeColor = Color.Green));
            label5.Invoke((MethodInvoker)(() => label5.ForeColor = Color.Green));
            label6.Invoke((MethodInvoker)(() => label6.ForeColor = Color.Green));
            try
            {
                if (checkBox12.Checked)
                {
                    if (File.Exists(textBox1.Text))
                    {
                        vba(textBox1.Text);
                    }
                    if (File.Exists(textBox2.Text))
                    {
                        vba(textBox2.Text);
                    }
                    if (File.Exists(textBox3.Text))
                    {
                        vba(textBox3.Text);
                    }
                }
                if (checkBox8.Checked)
                {
                    if (File.Exists(textBox1.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox1.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                    if (File.Exists(textBox2.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox2.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                    if (File.Exists(textBox3.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox3.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                }
                DataSet DsF = new DataSet();

                if (checkBox1.Checked == true)
                {
                    DsF.Tables.Add(ByProduct().ToTable());
                }
                if (checkBox2.Checked == true)
                {
                    DsF.Tables.Add(ByCustomer().ToTable());
                }
                if (checkBox3.Checked == true)
                {
                    DsF.Tables.Add(SCR().ToTable());
                }
                if (checkBox4.Checked == true)
                {
                    DsF.Tables.Add(ByArea().ToTable());
                }
                if (checkBox5.Checked == true)
                {
                    DsF.Tables.Add(AvgResolveTime().ToTable());
                }
                if (checkBox6.Checked == true)
                {
                    DsF.Tables.Add(ByAgeing().ToTable());
                }
                if (checkBox9.Checked == true)
                {
                    DsF.Tables.Add(EnvMetrics().ToTable());
                }
                if (checkBox10.Checked == true)
                {
                    DsF.Tables.Add(IncidentStatus().ToTable());
                }
                if (checkBox11.Checked == true)
                {
                    DsF.Tables.Add(LeakageRate().ToTable());
                }
                int attempt = 0;
            fl:
                if (File.Exists(Path.GetDirectoryName(path) + filename))
                {
                    filename = filename.Replace("(" + attempt + ")", "");
                    attempt++;
                    filename = filename.Replace(".", "(" + attempt + ").");
                    goto fl;
                    //diaRes = MessageBox.Show("File Already Exists...", "No Big Deal...", MessageBoxButtons.AbortRetryIgnore);
                    //if (diaRes.Equals(DialogResult.Ignore)) { }
                    //if (diaRes.Equals(DialogResult.Abort)) { throw new IOException(); }
                    //if (diaRes.Equals(DialogResult.Retry)) { goto fl; }
                }
                //CreateWorkbook(Path.GetDirectoryName(textBox1.Text) + "\\SOAReport.xls", DsF);                
                ExportDataTableToExcel(Path.GetDirectoryName(path) + filename, DsF, checkBox7.Checked, textBox4.Text);
                if (checkBox13.Checked == true)
                {
                    genCompactSheet(Path.GetDirectoryName(path) + filename, textBox2.Text);
                }
                if (MessageBox.Show("Report Generated : \u2714 \n" + Path.GetDirectoryName(path) + filename + "\n\nOpen it??", "Hey Look I have Done Everything", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                {
                    Process.Start(Path.GetDirectoryName(path) + filename);
                }
            }
            catch (IndexOutOfRangeException ex)
            {
                if (!File.Exists(textBox1.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label4.Text = "Main : File Missing";
                    //label4.ForeColor = Color.Red;
                    label4.Invoke((MethodInvoker)(() => label4.Text = "Main : File Missing"));
                    label4.Invoke((MethodInvoker)(() => label4.ForeColor = Color.Red));
                }
                else if (!File.Exists(textBox2.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label5.Text = "Resolved Sheet : File Missing";
                    //label5.ForeColor = Color.Red;
                    label5.Invoke((MethodInvoker)(() => label5.Text = "Resolved : File Missing"));
                    label5.Invoke((MethodInvoker)(() => label5.ForeColor = Color.Red));

                }
                else if (!File.Exists(textBox3.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label6.Text = "Age From Open Date : File Missing";
                    //label6.ForeColor = Color.Red;
                    label6.Invoke((MethodInvoker)(() => label6.Text = "Age From Open Date : File Missing"));
                    label6.Invoke((MethodInvoker)(() => label6.ForeColor = Color.Red));
                }
                else
                {
                    MessageBox.Show("Something Went Wrong...\n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com\n\n" + ex.Message, "Error Occured");
                }

            }
            catch (FileNotFoundException)
            {
                if (!File.Exists(textBox1.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label4.Text = "Main : File Missing";
                    //label4.ForeColor = Color.Red;
                    label4.Invoke((MethodInvoker)(() => label4.Text = "Main : File Missing"));
                    label4.Invoke((MethodInvoker)(() => label4.ForeColor = Color.Red));
                }
                else if (!File.Exists(textBox2.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label5.Text = "Resolved Sheet : File Missing";
                    //label5.ForeColor = Color.Red;
                    label5.Invoke((MethodInvoker)(() => label5.Text = "Resolved : File Missing"));
                    label5.Invoke((MethodInvoker)(() => label5.ForeColor = Color.Red));

                }
                else if (!File.Exists(textBox3.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label6.Text = "Age From Open Date : File Missing";
                    //label6.ForeColor = Color.Red;
                    label6.Invoke((MethodInvoker)(() => label6.Text = "Age From Open Date : File Missing"));
                    label6.Invoke((MethodInvoker)(() => label6.ForeColor = Color.Red));
                }
                else if (!File.Exists(Path.GetDirectoryName(path) + filename))
                {
                    MessageBox.Show("Step 1 output seems missing", "Catastrophy...");
                    //label6.Text = "Age From Open Date : File Missing";
                    //label6.ForeColor = Color.Red;
                    //label6.Invoke((MethodInvoker)(() => label6.Text = "Age From Open Date : File Missing"));
                    //label6.Invoke((MethodInvoker)(() => label6.ForeColor = Color.Red));
                }
            }
            catch (IOException)
            {
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("Excel Not in Proper Format:\n" + ex.Message + ", \n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com", "Error Occured");
            }
            catch (Exception ex)
            {
                if (!File.Exists(textBox1.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label4.Text = "Main : File Missing";
                    //label4.ForeColor = Color.Red;
                    label4.Invoke((MethodInvoker)(() => label4.Text = "Main : File Missing"));
                    label4.Invoke((MethodInvoker)(() => label4.ForeColor = Color.Red));
                }
                else if (!File.Exists(textBox2.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label5.Text = "Resolved Sheet : File Missing";
                    //label5.ForeColor = Color.Red;
                    label5.Invoke((MethodInvoker)(() => label5.Text = "Resolved : File Missing"));
                    label5.Invoke((MethodInvoker)(() => label5.ForeColor = Color.Red));

                }
                else if (!File.Exists(textBox3.Text))
                {
                    MessageBox.Show("Dude Select A Proper File First....", "No Big Deal...");
                    //label6.Text = "Age From Open Date : File Missing";
                    //label6.ForeColor = Color.Red;
                    label6.Invoke((MethodInvoker)(() => label6.Text = "Age From Open Date : File Missing"));
                    label6.Invoke((MethodInvoker)(() => label6.ForeColor = Color.Red));
                }
                else
                {
                    MessageBox.Show("Something Went Wrong...\n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com\n\n" + ex.Message, "Error Occured");
                }
            }
        Exit_Report: ;
        button2.Invoke((MethodInvoker)(() => button2.Enabled = true));
        }
        public static void CreateWorkbook(String filePath, DataSet dataset)
        {
            int minRows = 15;
            if (dataset.Tables.Count == 0)
                throw new ArgumentException("DataSet needs to have at least one DataTable", "dataset");

            Workbook workbook = new Workbook();
            foreach (DataTable dt in dataset.Tables)
            {
                Worksheet worksheet = new Worksheet(dt.TableName);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    worksheet.Cells[0, i] = new Cell(dt.Columns[i].ColumnName);
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        try
                        {
                            Convert.ToDateTime(dt.Rows[j][i]);
                            worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i], CellFormat.Date);
                        }
                        catch (FormatException)
                        {
                            worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                        }
                        catch (InvalidCastException)
                        {
                            worksheet.Cells[j + 1, i] = new Cell(dt.Rows[j][i] == DBNull.Value ? string.Empty : dt.Rows[j][i]);
                            //worksheet.Cells[j + 1, i] = new Cell("");
                        }
                    }
                }
                //For Excel 2010 Bug of small files by ExcelLibrary
                if (dt.Rows.Count < minRows)
                {
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {

                        for (int row = dt.Rows.Count + 1; row < minRows; row++)
                        {
                            worksheet.Cells[row, col] = new Cell(" ");
                        }
                    }
                }

                workbook.Worksheets.Add(worksheet);
            }
            workbook.Save(filePath);
        }

        public DataView AvgResolveTime()
        {
            //label13.Text = "...";
            //label13.BeginInvoke(delegate { label13.Text = "..."; });
            string query;
            label13.Invoke((MethodInvoker)(() => label13.Text = "..."));
            DataSet ds = new DataSet();
            DataSet ds_sub = new DataSet();
            OleDbDataAdapter adaptor;
            DataTable dtf = new DataTable("Average time to resolve");
            TimeSpan timediffB = new TimeSpan();
            TimeSpan timediffP1 = new TimeSpan();
            TimeSpan timediffP2 = new TimeSpan();
            TimeSpan timediffP3 = new TimeSpan();
            TimeSpan timediffP4 = new TimeSpan();
            TimeSpan timediffP5 = new TimeSpan();
            dtf.Columns.Add("Product Categorization Tier1");
            dtf.Columns.Add("Product Categorization Tier2");
            dtf.Columns.Add("P1 Average Time(Days)");
            dtf.Columns.Add("P2 Average Time(Days)");
            dtf.Columns.Add("P3 Average Time(Days)");
            dtf.Columns.Add("P4 Average Time(Days)");
            dtf.Columns.Add("P5 Average Time(Days)");
            dtf.Columns.Add("Blank Average Time(Days)");
            dtf.Columns.Add("Total Issues Resolved");
            if (!dateTimePicker1.Enabled)
            {
                if (!textBox2.Text.Equals(""))
                {
                    path = textBox2.Text;
                    query = "select * from [<placeHolder>]";// where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                }
                else
                {
                    path = textBox1.Text;
                    query = "select * from [<placeHolder>]";// where [Opened Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                }
            }
            else
            {
                if (!textBox2.Text.Equals(""))
                {
                    path = textBox2.Text;
                    query = "select * from [<placeHolder>] where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                }
                else
                {
                    path = textBox1.Text;
                    query = "select * from [<placeHolder>] where [Opened Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                }
            }
            int found = -1;
            ArrayList category1 = new ArrayList();
            ArrayList category2 = new ArrayList();
            //Console.WriteLine("Entered Path: " + path);
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                OleDbCommand command;
                connection.Open();
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    try
                    {
                        foreach (object cat in category1)
                        {
                            if (cat.Equals(dr["Product Categorization Tier1"]))
                                found++;
                        }
                        if (found == -1)
                        {
                            category1.Add(dr["Product Categorization Tier1"]);
                        }
                        else
                        {
                            found = -1;
                        }
                    }
                    catch (InvalidCastException)
                    {
                        if (dr["Product Categorization Tier1"] == null)
                            category1.Add(null);
                    }
                }
                DataTable dt = ds.Tables[0].Clone();
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                try
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(DBNull.Value))
                        {
                            dra = ds.Tables[0].Select("[Product Categorization Tier1] IS NULL");
                        }
                        else
                            dra = ds.Tables[0].Select("[Product Categorization Tier1]='" + cat + "'");
                        //dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + cat + "'");//.CopyTo(drc,0);
                        foreach (DataRow drt in dra)
                        {
                            dt.ImportRow(drt);
                            try
                            {
                                foreach (object cat1 in category2)
                                {
                                    if (cat1.Equals(drt["Product Categorization Tier2"]))
                                        found++;
                                }
                                if (found == -1)
                                {
                                    category2.Add(drt["Product Categorization Tier2"]);
                                }
                                else
                                {
                                    found = -1;
                                }
                            }
                            catch (InvalidCastException)
                            {
                                if (drt["Product Categorization Tier2"] == null)
                                    category2.Add(null);
                            }
                        }
                        DataRow[] dra1;
                        int p1 = 0, p2 = 0, p3 = 0, p4 = 0, p5 = 0, pb = 0;
                        int pr1 = 0, pr2 = 0, pr3 = 0, pr4 = 0, pr5 = 0, prb = 0;
                        int prd1 = 0, prd2 = 0, prd3 = 0, prd4 = 0, prd5 = 0, prdb = 0;
                        foreach (object cat2 in category2)
                        {
                            if (cat2.Equals(DBNull.Value))
                            {
                                dra1 = dt.Select("[Product Categorization Tier2] IS NULL");
                            }
                            else
                                dra1 = dt.Select("[Product Categorization Tier2]='" + cat2 + "'");
                            foreach (DataRow dr in dra1)
                            {
                                if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                                {
                                    pb++;
                                    if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                    {
                                        prb++;
                                        try
                                        {
                                            timediffB += (Convert.ToDateTime(dr["Resolved Date"])).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                            prdb++;
                                        }
                                        catch (InvalidCastException) { }
                                    }
                                }
                                else
                                {
                                    if (Convert.ToInt32(dr["Priority"]).Equals(1))
                                    {
                                        p1++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr1++;
                                            try
                                            {
                                                timediffP1 += (Convert.ToDateTime(dr["Resolved Date"])).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                                prd1++;
                                            }
                                            catch (InvalidCastException) { }
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(2))
                                    {
                                        p2++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr2++;
                                            try
                                            {
                                                timediffP2 += (Convert.ToDateTime(dr["Resolved Date"])).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                                prd2++;
                                            }
                                            catch (InvalidCastException) { }
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(3))
                                    {
                                        p3++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr3++;
                                            try
                                            {
                                                timediffP3 += (Convert.ToDateTime(dr["Resolved Date"])).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                                prd3++;
                                            }
                                            catch (InvalidCastException) { }
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(4))
                                    {
                                        p4++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr4++;
                                            try
                                            {
                                                timediffP4 += (Convert.ToDateTime(dr["Resolved Date"])).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                                prd4++;
                                            }
                                            catch (InvalidCastException) { }
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(5))
                                    {
                                        p5++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr5++;
                                            try
                                            {
                                                timediffP5 += (Convert.ToDateTime(dr["Resolved Date"])).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                                prd5++;
                                            }
                                            catch (InvalidCastException) { }
                                        }
                                    }
                                }

                            }
                            DataRow drf = dtf.NewRow();
                            drf[0] = dra1[0]["Product Categorization Tier1"];
                            drf[1] = dra1[0]["Product Categorization Tier2"];
                            drf[2] = p1;
                            drf[3] = p2;
                            drf[4] = p3;
                            drf[5] = p4;
                            drf[6] = p5;
                            drf[7] = pb;
                            //drf[8] = p1 + p2 + p3 + p4 + p5 + pb;
                            //drf[9] = pr1;
                            //drf[10] = pr2;
                            //drf[11] = pr3;
                            //drf[12] = pr4;
                            //drf[13] = pr5;
                            //drf[14] = prb;
                            drf[8] = pr1 + pr2 + pr3 + pr4 + pr5 + prb;
                            decimal avPB = 0, avP1 = 0, avP2 = 0, avP3 = 0, avP4 = 0, avP5 = 0;
                            if (timediffB.Days != 0)
                                avPB = Convert.ToDecimal(timediffB.Days) / prdb;
                            if (timediffP1.Days != 0)
                                avP1 = Convert.ToDecimal(timediffP1.Days) / prd1;
                            if (timediffP2.Days != 0)
                                avP2 = Convert.ToDecimal(timediffP2.Days) / prd2;
                            if (timediffP3.Days != 0)
                                avP3 = Convert.ToDecimal(timediffP3.Days) / prd3;
                            if (timediffP4.Days != 0)
                                avP4 = Convert.ToDecimal(timediffP4.Days) / prd4;
                            if (timediffP5.Days != 0)
                                avP5 = Convert.ToDecimal(timediffP5.Days) / prd5;
                            drf[2] = Math.Round(avP1, 1);
                            drf[3] = Math.Round(avP2, 1);
                            drf[4] = Math.Round(avP3, 1);
                            drf[5] = Math.Round(avP4, 1);
                            drf[6] = Math.Round(avP5, 1);
                            drf[7] = Math.Round(avPB, 1);
                            p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = prd1 = prd2 = prd3 = prd4 = prd5 = prdb = 0;
                            timediffB = timediffP1 = timediffP2 = timediffP3 = timediffP4 = timediffP5 = TimeSpan.Zero;
                            dtf.Rows.Add(drf);
                        }
                        dt.Clear();
                        category2.Clear();
                    }
                }
                catch (InvalidCastException)
                {
                    dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + null + "'");
                }
            }
            //label13.Text = "\u2714";
            label13.Invoke((MethodInvoker)(() => label13.Text = "\u2714"));
            DataView dv_temp;
            dv_temp = new DataView(dtf);
            dv_temp.Sort = "[Product Categorization Tier1] desc";
            return dv_temp;
        }
        public DataView ByAgeing()
        {
            string query;
            //label14.Text = "...";
            label14.Invoke((MethodInvoker)(() => label14.Text = "..."));
            DataSet ds = new DataSet();
            DataSet ds_sub = new DataSet();
            OleDbDataAdapter adaptor;
            DataTable dtf = new DataTable("Age From Open Date");
            dtf.Columns.Add("Product Categorization Tier1");
            dtf.Columns.Add("Product Categorization Tier2");
            dtf.Columns.Add("P1_30");
            dtf.Columns.Add("P2_30");
            dtf.Columns.Add("P3_30");
            dtf.Columns.Add("P4_30");
            dtf.Columns.Add("P5_30");
            dtf.Columns.Add("PB_30");
            dtf.Columns.Add("P1_15");
            dtf.Columns.Add("P2_15");
            dtf.Columns.Add("P3_15");
            dtf.Columns.Add("P4_15");
            dtf.Columns.Add("P5_15");
            dtf.Columns.Add("PB_15");
            dtf.Columns.Add("P1_7");
            dtf.Columns.Add("P2_7");
            dtf.Columns.Add("P3_7");
            dtf.Columns.Add("P4_7");
            dtf.Columns.Add("P5_7");
            dtf.Columns.Add("PB_7");
            dtf.Columns.Add("P1_3");
            dtf.Columns.Add("P2_3");
            dtf.Columns.Add("P3_3");
            dtf.Columns.Add("P4_3");
            dtf.Columns.Add("P5_3");
            dtf.Columns.Add("PB_3");
            dtf.Columns.Add("P1");
            dtf.Columns.Add("P2");
            dtf.Columns.Add("P3");
            dtf.Columns.Add("P4");
            dtf.Columns.Add("P5");
            dtf.Columns.Add("PB");
            dtf.Columns.Add("Total Issues");
            if (!textBox3.Text.Equals(""))
                path = textBox3.Text;
            else
            {
                MessageBox.Show("Open Not Resolved Sheet not present Going by Opened Instead.");
                path = textBox1.Text;
            }
            if (!dateTimePicker1.Enabled)
            {
                query = "select * from [<placeHolder>]";// where [Opened Date] < #" + dateTimePicker2.Value + "#";
            }
            else
            {
                query = "select * from [<placeHolder>] where [Opened Date] < #" + dateTimePicker2.Value + "#";
            }
            int found = -1;
            ArrayList category1 = new ArrayList();
            ArrayList category2 = new ArrayList();
            int p1, p2, p3, p4, p5, pb, pb_30, p1_30, p2_30, p3_30, p4_30, p5_30, p1_15, p2_15, p3_15, p4_15, p5_15, pb_15, p1_7, p2_7, p3_7, p4_7, p5_7, pb_7, p1_3, p2_3, p3_3, p4_3, p5_3, pb_3;
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                OleDbCommand command;
                connection.Open();
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    try
                    {
                        foreach (object cat in category1)
                        {
                            if (cat.Equals(dr["Product Categorization Tier1"]))
                                found++;
                        }
                        if (found == -1)
                        {
                            category1.Add(dr["Product Categorization Tier1"]);
                        }
                        else
                        {
                            found = -1;
                        }
                    }
                    catch (InvalidCastException)
                    {
                        if (dr["Product Categorization Tier1"] == null)
                            category1.Add(null);
                    }
                }
                DataTable dt = ds.Tables[0].Clone();
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                try
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(DBNull.Value))
                        {
                            dra = ds.Tables[0].Select("[Product Categorization Tier1] IS NULL");
                        }
                        else
                            dra = ds.Tables[0].Select("[Product Categorization Tier1]='" + cat + "'");
                        //dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + cat + "'");//.CopyTo(drc,0);
                        foreach (DataRow drt in dra)
                        {
                            dt.ImportRow(drt);
                            try
                            {
                                foreach (object cat1 in category2)
                                {
                                    if (cat1.Equals(drt["Product Categorization Tier2"]))
                                        found++;
                                }
                                if (found == -1)
                                {
                                    category2.Add(drt["Product Categorization Tier2"]);
                                }
                                else
                                {
                                    found = -1;
                                }
                            }
                            catch (InvalidCastException)
                            {
                                if (drt["Product Categorization Tier2"] == null)
                                    category2.Add(null);
                            }
                        }
                        DataRow[] dra1;
                        TimeSpan timediff = TimeSpan.Zero;
                        foreach (object cat2 in category2)
                        {
                            p1 = p2 = p3 = p4 = p5 = pb = pb_30 = p1_30 = p2_30 = p3_30 = p4_30 = p5_30 = p1_15 = p2_15 = p3_15 = p4_15 = p5_15 = pb_15 = p1_7 = p2_7 = p3_7 = p4_7 = p5_7 = pb_7 = p1_3 = p2_3 = p3_3 = p4_3 = p5_3 = pb_3 = 0;
                            if (cat2.Equals(DBNull.Value))
                            {
                                dra1 = dt.Select("[Product Categorization Tier2] IS NULL");
                            }
                            else
                                dra1 = dt.Select("[Product Categorization Tier2]='" + cat2 + "'");
                            foreach (DataRow dr in dra1)
                            {
                                if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                                {
                                    try
                                    {
                                        timediff = (DateTime.Now).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                    }
                                    catch (InvalidCastException) { }
                                    if (timediff.Days > 30)
                                        pb_30++;
                                    else if (timediff.Days > 15)
                                        pb_15++;
                                    else if (timediff.Days > 7)
                                        pb_7++;
                                    else if (timediff.Days > 3)
                                        pb_3++;
                                    else if (timediff.Days > 1)
                                        pb++;
                                }
                                else
                                {
                                    if (Convert.ToInt32(dr["Priority"]).Equals(1))
                                    {
                                        try
                                        {
                                            timediff = (DateTime.Now).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                        }
                                        catch (InvalidCastException) { }
                                        if (timediff.Days > 30)
                                            p1_30++;
                                        else if (timediff.Days > 15)
                                            p1_15++;
                                        else if (timediff.Days > 7)
                                            p1_7++;
                                        else if (timediff.Days > 3)
                                            p1_3++;
                                        else if (timediff.Days > 1)
                                            p1++;
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(2))
                                    {
                                        try
                                        {
                                            timediff = (DateTime.Now).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                        }
                                        catch (InvalidCastException) { }
                                        if (timediff.Days > 30)
                                            p2_30++;
                                        else if (timediff.Days > 15)
                                            p2_15++;
                                        else if (timediff.Days > 7)
                                            p2_7++;
                                        else if (timediff.Days > 3)
                                            p2_3++;
                                        else if (timediff.Days > 1)
                                            p2++;
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(3))
                                    {
                                        try
                                        {
                                            timediff = (DateTime.Now).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                        }
                                        catch (InvalidCastException) { }
                                        if (timediff.Days > 30)
                                            p3_30++;
                                        else if (timediff.Days > 15)
                                            p3_15++;
                                        else if (timediff.Days > 7)
                                            p3_7++;
                                        else if (timediff.Days > 3)
                                            p3_3++;
                                        else if (timediff.Days > 1)
                                            p3++;
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(4))
                                    {
                                        try
                                        {
                                            timediff = (DateTime.Now).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                        }
                                        catch (InvalidCastException) { }
                                        if (timediff.Days > 30)
                                            p4_30++;
                                        else if (timediff.Days > 15)
                                            p4_15++;
                                        else if (timediff.Days > 7)
                                            p4_7++;
                                        else if (timediff.Days > 3)
                                            p4_3++;
                                        else if (timediff.Days > 1)
                                            p4++;
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(5))
                                    {
                                        try
                                        {
                                            timediff = (DateTime.Now).Date - (Convert.ToDateTime(dr["Opened Date"])).Date;
                                        }
                                        catch (InvalidCastException) { }
                                        if (timediff.Days > 30)
                                            p5_30++;
                                        else if (timediff.Days > 15)
                                            p5_15++;
                                        else if (timediff.Days > 7)
                                            p5_7++;
                                        else if (timediff.Days > 3)
                                            p5_3++;
                                        else if (timediff.Days > 1)
                                            p5++;
                                    }
                                }

                            }
                            DataRow drf = dtf.NewRow();
                            drf[0] = dra1[0]["Product Categorization Tier1"];
                            drf[1] = dra1[0]["Product Categorization Tier2"];
                            drf[2] = p1_30;
                            drf[3] = p2_30;
                            drf[4] = p3_30;
                            drf[5] = p4_30;
                            drf[6] = p5_30;
                            drf[7] = pb_30;
                            drf[8] = p1_15;
                            drf[9] = p2_15;
                            drf[10] = p3_15;
                            drf[11] = p4_15;
                            drf[12] = p5_15;
                            drf[13] = pb_15;
                            drf[14] = p1_7;
                            drf[15] = p2_7;
                            drf[16] = p3_7;
                            drf[17] = p4_7;
                            drf[18] = p5_7;
                            drf[19] = pb_7;
                            drf[20] = p1_3;
                            drf[21] = p2_3;
                            drf[22] = p3_3;
                            drf[23] = p4_3;
                            drf[24] = p5_3;
                            drf[25] = pb_3;
                            drf[26] = p1;
                            drf[27] = p2;
                            drf[28] = p3;
                            drf[29] = p4;
                            drf[30] = p5;
                            drf[31] = pb;
                            drf[32] = p1_30 + p2_30 + p3_30 + p4_30 + p5_30 + pb_30 + p1_15 + p2_15 + p3_15 + p4_15 + p5_15 + pb_15 + p1_7 + p2_7 + p3_7 + p4_7 + p5_7 + pb_7 + p1_3 + p2_3 + p3_3 + p4_3 + p5_3 + pb_3 + p1 + p2 + p3 + p4 + p5 + pb;
                            //p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = 0;
                            dtf.Rows.Add(drf);
                        }
                        dt.Clear();
                        category2.Clear();
                    }
                }
                catch (InvalidCastException)
                {
                    dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + null + "'");
                }
            }
            //label14.Text = "\u2714";
            label14.Invoke((MethodInvoker)(() => label14.Text = "\u2714"));
            DataView dv_temp = new DataView(dtf);
            dv_temp.Sort = "[Product Categorization Tier1] desc, [Product Categorization Tier2] asc";
            return dv_temp;
        }
        public DataView ByArea()
        {
            string query = "select * from [<placeHolder>] where [Opened Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            string query1 = "select * from [<placeHolder>] where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            if (!dateTimePicker1.Enabled)
            {
                query = "select * from [<placeHolder>]";// where [Opened Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                query1 = "select * from [<placeHolder>]";// where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            }
            //label12.Text = "...";
            label12.Invoke((MethodInvoker)(() => label12.Text = "..."));
            string path1;
            DataView dv_temp;
            OleDbDataAdapter adaptor;
            DataTable dtf = new DataTable("By Region");
            dtf.Columns.Add("Region/Area");
            dtf.Columns.Add("P1");
            dtf.Columns.Add("P2");
            dtf.Columns.Add("P3");
            dtf.Columns.Add("P4");
            dtf.Columns.Add("P5");
            dtf.Columns.Add("Blank");
            dtf.Columns.Add("Total Issues");
            dtf.Columns.Add("P1R");
            dtf.Columns.Add("P2R");
            dtf.Columns.Add("P3R");
            dtf.Columns.Add("P4R");
            dtf.Columns.Add("P5R");
            dtf.Columns.Add("BlankR");
            dtf.Columns.Add("Issues Resolved");
            dtf.Columns.Add("% Resolved");
            DataTable dtf1 = new DataTable("By Region");
            dtf1.Columns.Add("Region/Area");
            dtf1.Columns.Add("P1");
            dtf1.Columns.Add("P2");
            dtf1.Columns.Add("P3");
            dtf1.Columns.Add("P4");
            dtf1.Columns.Add("P5");
            dtf1.Columns.Add("Blank");
            dtf1.Columns.Add("Total Issues");
            dtf1.Columns.Add("P1R");
            dtf1.Columns.Add("P2R");
            dtf1.Columns.Add("P3R");
            dtf1.Columns.Add("P4R");
            dtf1.Columns.Add("P5R");
            dtf1.Columns.Add("BlankR");
            dtf1.Columns.Add("Issues Resolved");
            dtf1.Columns.Add("% Resolved");
            path = textBox1.Text;
            if (textBox2.Text.Equals(""))
            {
                if (MessageBox.Show("Continue Region/Area without Resolved Sheet", "Resolved Sheet not selected: Region/Area", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                {
                    path1 = textBox1.Text;
                }
                else
                {
                    path1 = "";
                    throw new FileNotFoundException();
                }
            }
            else
            {
                path1 = textBox2.Text;
            }
            int found = -1;
            ArrayList category1 = new ArrayList();
            //Console.WriteLine("Entered Path: " + path);
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                DataSet ds = new DataSet();
                connection.Open();
                OleDbCommand command;
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                //OleDbCommand command = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "]", connection);
                //OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(dr["Region/Area"]))
                            found++;
                    }
                    if (found == -1)
                    {
                        category1.Add(dr["Region/Area"]);
                    }
                    else
                    {
                        found = -1;
                    }

                }
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                foreach (object cat in category1)
                {
                    if (cat.Equals(DBNull.Value))
                    {
                        dra = ds.Tables[0].Select("[Region/Area] IS NULL");
                    }
                    else
                        dra = ds.Tables[0].Select("[Region/Area]='" + cat + "'");


                    int p1 = 0, p2 = 0, p3 = 0, p4 = 0, p5 = 0, pb = 0;
                    int pr1 = 0, pr2 = 0, pr3 = 0, pr4 = 0, pr5 = 0, prb = 0;
                    foreach (DataRow dr in dra)
                    {
                        if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                        {
                            pb++;
                            if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                            {
                                prb++;
                            }
                        }
                        else
                        {
                            if (Convert.ToInt32(dr["Priority"]).Equals(1))
                            {
                                p1++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr1++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(2))
                            {
                                p2++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr2++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(3))
                            {
                                p3++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr3++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(4))
                            {
                                p4++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr4++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(5))
                            {
                                p5++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr5++;
                                }
                            }
                        }
                    }

                    DataRow drf = dtf.NewRow();

                    drf[0] = dra[0]["Region/Area"];
                    drf[1] = p1;
                    drf[2] = p2;
                    drf[3] = p3;
                    drf[4] = p4;
                    drf[5] = p5;
                    drf[6] = pb;
                    drf[7] = p1 + p2 + p3 + p4 + p5 + pb;
                    drf[8] = 0;// pr1;
                    drf[9] = 0;//pr2;
                    drf[10] = 0;//pr3;
                    drf[11] = 0;//pr4;
                    drf[12] = 0;//pr5;
                    drf[13] = 0;// prb;
                    drf[14] = 0;//pr1 + pr2 + pr3 + pr4 + pr5 + prb;
                    p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = 0;
                    dtf.Rows.Add(drf);
                }
            }
            found = -1;
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                DataSet ds = new DataSet();
                connection.Open();
                OleDbCommand command;
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query1.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query1.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(dr["Region/Area"]))
                            found++;
                    }
                    if (found == -1)
                    {
                        category1.Add(dr["Region/Area"]);
                    }
                    else
                    {
                        found = -1;
                    }

                }
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                foreach (object cat in category1)
                {
                    if (cat.Equals(DBNull.Value))
                    {
                        dra = ds.Tables[0].Select("[Region/Area] IS NULL");
                    }
                    else
                        dra = ds.Tables[0].Select("[Region/Area]='" + cat + "'");


                    int p1 = 0, p2 = 0, p3 = 0, p4 = 0, p5 = 0, pb = 0;
                    int pr1 = 0, pr2 = 0, pr3 = 0, pr4 = 0, pr5 = 0, prb = 0;
                    foreach (DataRow dr in dra)
                    {
                        if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                        {
                            pb++;
                            if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                            {
                                prb++;
                            }
                        }
                        else
                        {
                            if (Convert.ToInt32(dr["Priority"]).Equals(1))
                            {
                                p1++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr1++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(2))
                            {
                                p2++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr2++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(3))
                            {
                                p3++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr3++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(4))
                            {
                                p4++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr4++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(5))
                            {
                                p5++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr5++;
                                }
                            }
                        }
                    }

                    DataRow drf = dtf1.NewRow();
                    try
                    {
                        drf[0] = cat;
                        drf[1] = 0;// p1;
                        drf[2] = 0;// p2;
                        drf[3] = 0;//p3;
                        drf[4] = 0;//p4;
                        drf[5] = 0;//p5;
                        drf[6] = 0;//pb;
                        drf[7] = 0;//p1 + p2 + p3 + p4 + p5 + pb;
                        drf[8] = pr1;
                        drf[9] = pr2;
                        drf[10] = pr3;
                        drf[11] = pr4;
                        drf[12] = pr5;
                        drf[13] = prb;
                        drf[14] = pr1 + pr2 + pr3 + pr4 + pr5 + prb;
                        p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = 0;
                        dtf1.Rows.Add(drf);
                    }
                    catch (IndexOutOfRangeException)
                    {

                    }
                }
            }
            List<string> regionList = new List<string>();
            foreach (DataRow drf in dtf1.Rows)
            {
                foreach (DataRow drT in dtf.Rows)
                {
                    if (drf["Region/Area"].Equals(drT["Region/Area"]))
                    {
                        drf["P1"] = drT["P1"];
                        drf["P2"] = drT["P2"];
                        drf["P3"] = drT["P3"];
                        drf["P4"] = drT["P4"];
                        drf["P5"] = drT["P5"];
                        drf["Blank"] = drT["Blank"];
                        drf["Total Issues"] = Convert.ToInt32(drT["P1"]) + Convert.ToInt32(drT["P2"]) + Convert.ToInt32(drT["P3"]) + Convert.ToInt32(drT["P4"]) + Convert.ToInt32(drT["P5"]) + Convert.ToInt32(drT["Blank"]);
                        regionList.Add(drT["Region/Area"].ToString());
                    }
                }
            }
            foreach (DataRow drT in dtf.Rows)
            {
                if (!regionList.Contains(drT["Region/Area"].ToString())/* && !catrgory_1.Contains(drT["Product Categorization Tier1"].ToString())*/)
                {
                    dtf1.ImportRow(drT);
                }
            }
            //label12.Text = "\u2714";
            label12.Invoke((MethodInvoker)(() => label12.Text = "\u2714"));
            dv_temp = new DataView(dtf1);
            dv_temp.Sort = "Region/Area desc";
            return dv_temp;
        }   
        public DataView ByProduct()
        {
            string query = "select * from [<placeHolder>] where [Opened Date] between  #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            string query1 = "select * from [<placeHolder>] where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            if (!dateTimePicker1.Enabled)
            {
                query = "select * from [<placeHolder>]";// where [Opened Date] between  #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                query1 = "select * from [<placeHolder>]";// where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            }
            label9.Invoke((MethodInvoker)(() => label9.Text = "..."));
            string path1;
            OleDbDataAdapter adaptor;
            DataTable dtf = new DataTable("By Product");
            dtf.Columns.Add("Product Categorization Tier1");
            dtf.Columns.Add("Product Categorization Tier2");
            dtf.Columns.Add("P1");
            dtf.Columns.Add("P2");
            dtf.Columns.Add("P3");
            dtf.Columns.Add("P4");
            dtf.Columns.Add("P5");
            dtf.Columns.Add("Blank");
            dtf.Columns.Add("Total Issues");
            dtf.Columns.Add("P1R");
            dtf.Columns.Add("P2R");
            dtf.Columns.Add("P3R");
            dtf.Columns.Add("P4R");
            dtf.Columns.Add("P5R");
            dtf.Columns.Add("BlankR");
            dtf.Columns.Add("Issues Resolved");
            dtf.Columns.Add("% Resolved");
            DataTable dtf1 = new DataTable("By Product");
            dtf1.Columns.Add("Product Categorization Tier1");
            dtf1.Columns.Add("Product Categorization Tier2");
            dtf1.Columns.Add("P1");
            dtf1.Columns.Add("P2");
            dtf1.Columns.Add("P3");
            dtf1.Columns.Add("P4");
            dtf1.Columns.Add("P5");
            dtf1.Columns.Add("Blank");
            dtf1.Columns.Add("Total Issues");
            dtf1.Columns.Add("P1R");
            dtf1.Columns.Add("P2R");
            dtf1.Columns.Add("P3R");
            dtf1.Columns.Add("P4R");
            dtf1.Columns.Add("P5R");
            dtf1.Columns.Add("BlankR");
            dtf1.Columns.Add("Issues Resolved");
            dtf1.Columns.Add("% Resolved");
            path = textBox1.Text;
            if (textBox2.Text.Equals(""))
            {
                if (MessageBox.Show("Continue Product without Resolved Sheet", "Resolved Sheet not selected: Product", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                {
                    path1 = textBox1.Text;
                }
                else
                {
                    path1 = "";
                    throw new FileNotFoundException();
                }
            }
            else
            {
                path1 = textBox2.Text;
            }
            int found = -1;
            ArrayList category1 = new ArrayList();
            ArrayList category2 = new ArrayList();
            //Console.WriteLine("Entered Path: " + path);
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES';"/*con*/))
            {
                DataSet ds = new DataSet();
                DataSet ds_sub = new DataSet();
                OleDbCommand command;
                connection.Open();
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    try
                    {
                        foreach (object cat in category1)
                        {
                            if (cat.Equals(dr["Product Categorization Tier1"]))
                                found++;
                        }
                        if (found == -1)
                        {
                            category1.Add(dr["Product Categorization Tier1"]);
                        }
                        else
                        {
                            found = -1;
                        }
                    }
                    catch (InvalidCastException)
                    {
                        if (dr["Product Categorization Tier1"] == null)
                            category1.Add(null);
                    }
                }
                DataTable dt = ds.Tables[0].Clone();
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                try
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(DBNull.Value))
                        {
                            dra = ds.Tables[0].Select("[Product Categorization Tier1] IS NULL");
                        }
                        else
                            dra = ds.Tables[0].Select("[Product Categorization Tier1]='" + cat + "'");
                        //dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + cat + "'");//.CopyTo(drc,0);
                        foreach (DataRow drt in dra)
                        {
                            dt.ImportRow(drt);
                            try
                            {
                                foreach (object cat1 in category2)
                                {
                                    if (cat1.Equals(drt["Product Categorization Tier2"]))
                                        found++;
                                }
                                if (found == -1)
                                {
                                    category2.Add(drt["Product Categorization Tier2"]);
                                }
                                else
                                {
                                    found = -1;
                                }
                            }
                            catch (InvalidCastException)
                            {
                                if (drt["Product Categorization Tier2"] == null)
                                    category2.Add(null);
                            }
                        }
                        DataRow[] dra1;
                        int p1 = 0, p2 = 0, p3 = 0, p4 = 0, p5 = 0, pb = 0;
                        int pr1 = 0, pr2 = 0, pr3 = 0, pr4 = 0, pr5 = 0, prb = 0;
                        foreach (object cat2 in category2)
                        {
                            if (cat2.Equals(DBNull.Value))
                            {
                                dra1 = dt.Select("[Product Categorization Tier2] IS NULL");
                            }
                            else
                                dra1 = dt.Select("[Product Categorization Tier2]='" + cat2 + "'");
                            foreach (DataRow dr in dra1)
                            {
                                if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                                {
                                    pb++;
                                    if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                    {
                                        prb++;
                                    }
                                }
                                else
                                {
                                    if (Convert.ToInt32(dr["Priority"]).Equals(1))
                                    {
                                        p1++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr1++;
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(2))
                                    {
                                        p2++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr2++;
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(3))
                                    {
                                        p3++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr3++;
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(4))
                                    {
                                        p4++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr4++;
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(5))
                                    {
                                        p5++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr5++;
                                        }
                                    }
                                }
                            }
                            DataRow drf = dtf.NewRow();
                            drf[0] = dra1[0]["Product Categorization Tier1"];
                            drf[1] = dra1[0]["Product Categorization Tier2"];
                            drf[2] = p1;
                            drf[3] = p2;
                            drf[4] = p3;
                            drf[5] = p4;
                            drf[6] = p5;
                            drf[7] = pb;
                            drf[8] = p1 + p2 + p3 + p4 + p5 + pb;
                            drf[9] = 0;//pr1;
                            drf[10] = 0;// pr2;
                            drf[11] = 0;// pr3;
                            drf[12] = 0;// pr4;
                            drf[13] = 0;// pr5;
                            drf[14] = 0;//prb;
                            drf[15] = 0;//pr1 + pr2 + pr3 + pr4 + pr5 + prb;
                            p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = 0;
                            dtf.Rows.Add(drf);
                        }
                        dt.Clear();
                        category2.Clear();
                    }
                }
                catch (InvalidCastException)
                {
                    dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + null + "'");
                }
            }
            found = -1;
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                DataSet ds = new DataSet();
                DataSet ds_sub = new DataSet();
                OleDbCommand command;
                connection.Open();
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query1.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query1.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    try
                    {
                        foreach (object cat in category1)
                        {
                            if (cat.Equals(dr["Product Categorization Tier1"]))
                                found++;
                        }
                        if (found == -1)
                        {
                            category1.Add(dr["Product Categorization Tier1"]);
                        }
                        else
                        {
                            found = -1;
                        }
                    }
                    catch (InvalidCastException)
                    {
                        if (dr["Product Categorization Tier1"] == null)
                            category1.Add(null);
                    }
                }
                DataTable dt = ds.Tables[0].Clone();
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                try
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(DBNull.Value))
                        {
                            dra = ds.Tables[0].Select("[Product Categorization Tier1] IS NULL");
                        }
                        else
                            dra = ds.Tables[0].Select("[Product Categorization Tier1]='" + cat + "'");
                        //dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + cat + "'");//.CopyTo(drc,0);
                        //if(dra
                        foreach (DataRow drt in dra)
                        {
                            dt.ImportRow(drt);
                            try
                            {
                                foreach (object cat1 in category2)
                                {
                                    if (cat1.Equals(drt["Product Categorization Tier2"]))
                                        found++;
                                }
                                if (found == -1)
                                {
                                    category2.Add(drt["Product Categorization Tier2"]);
                                }
                                else
                                {
                                    found = -1;
                                }
                            }
                            catch (InvalidCastException)
                            {
                                if (drt["Product Categorization Tier2"] == null)
                                    category2.Add(null);
                            }
                        }
                        DataRow[] dra1;
                        int p1 = 0, p2 = 0, p3 = 0, p4 = 0, p5 = 0, pb = 0;
                        int pr1 = 0, pr2 = 0, pr3 = 0, pr4 = 0, pr5 = 0, prb = 0;
                        foreach (object cat2 in category2)
                        {
                            if (cat2.Equals(DBNull.Value))
                            {
                                dra1 = dt.Select("[Product Categorization Tier2] IS NULL");
                            }
                            else
                                dra1 = dt.Select("[Product Categorization Tier2]='" + cat2 + "'");
                            foreach (DataRow dr in dra1)
                            {
                                if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                                {
                                    pb++;
                                    if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                    {
                                        prb++;
                                    }
                                }
                                else
                                {
                                    if (Convert.ToInt32(dr["Priority"]).Equals(1))
                                    {
                                        p1++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr1++;
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(2))
                                    {
                                        p2++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr2++;
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(3))
                                    {
                                        p3++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr3++;
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(4))
                                    {
                                        p4++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr4++;
                                        }
                                    }
                                    if (Convert.ToInt32(dr["Priority"]).Equals(5))
                                    {
                                        p5++;
                                        if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                        {
                                            pr5++;
                                        }
                                    }
                                }

                            }
                            DataRow drf = dtf1.NewRow();
                            drf[0] = cat;//dra1[0]["Product Categorization Tier1"];
                            drf[1] = cat2;//dra1[0]["Product Categorization Tier2"];
                            drf[2] = 0;//p1;
                            drf[3] = 0;//p2;
                            drf[4] = 0;//p3;
                            drf[5] = 0;// p4;
                            drf[6] = 0;//p5;
                            drf[7] = 0;// pb;
                            drf[8] = 0;//p1 + p2 + p3 + p4 + p5 + pb;
                            drf[9] = pr1;
                            drf[10] = pr2;
                            drf[11] = pr3;
                            drf[12] = pr4;
                            drf[13] = pr5;
                            drf[14] = prb;
                            drf[15] = pr1 + pr2 + pr3 + pr4 + pr5 + prb;
                            p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = 0;
                            dtf1.Rows.Add(drf);
                        }
                        dt.Clear();
                        category2.Clear();
                    }
                }
                catch (InvalidCastException)
                {
                    dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + null + "'");
                }
            }
            //Merging Assigned And Resolved.
            List<string> catrgory_1 = new List<string>();
            List<string> category_2 = new List<string>();
            //DataTable dtfM = new DataTable();
            foreach (DataRow drf in dtf1.Rows)
            {
                foreach (DataRow drT in dtf.Rows)
                {
                    if (drf["Product Categorization Tier1"].Equals(drT["Product Categorization Tier1"]) && drf["Product Categorization Tier2"].Equals(drT["Product Categorization Tier2"]))
                    {       //25 times
                        drf["P1"] = drT["P1"];
                        drf["P2"] = drT["P2"];
                        drf["P3"] = drT["P3"];
                        drf["P4"] = drT["P4"];
                        drf["P5"] = drT["P5"];
                        drf["Blank"] = drT["Blank"];
                        drf["Total Issues"] = Convert.ToInt32(drT["P1"]) + Convert.ToInt32(drT["P2"]) + Convert.ToInt32(drT["P3"]) + Convert.ToInt32(drT["P4"]) + Convert.ToInt32(drT["P5"]) + Convert.ToInt32(drT["Blank"]);
                        category_2.Add(drT["Product Categorization Tier2"].ToString());
                        catrgory_1.Add(drT["Product Categorization Tier1"].ToString());
                    }
                }
            }
            foreach (DataRow drT in dtf.Rows)
            {

                if (!category_2.Contains(drT["Product Categorization Tier2"].ToString()))///* && !catrgory_1.Contains(drT["Product Categorization Tier1"].ToString())*/)
                {
                    dtf1.ImportRow(drT);
                }
                else
                {
                    if (catrgory_1.Contains(drT["Product Categorization Tier1"].ToString())) { }
                    else
                    {
                        dtf1.ImportRow(drT);
                    }
                }
            }
            //label9.Text = "\u2714";
            label9.Invoke((MethodInvoker)(() => label9.Text = "\u2714"));
            DataView dv_temp;
            dv_temp = new DataView(dtf1);
            dv_temp.Sort = "[Product Categorization Tier1] desc";
            return dv_temp;
        }
        public DataView SCR()
        {
            string query;
            label11.Invoke((MethodInvoker)(() => label11.Text = "..."));
            DataSet ds = new DataSet();
            OleDbDataAdapter adaptor;
            DataView dv_temp;
            DataTable dtf = new DataTable("SCR Report");
            dtf.Columns.Add("Incident: Number");
            dtf.Columns.Add("Resolved Date");
            dtf.Columns.Add("Product Categorization Tier1");
            dtf.Columns.Add("Product Categorization Tier2");
            dtf.Columns.Add("SCR Details").DataType = System.Type.GetType("System.Double"); ;
            if (textBox2.Text.Equals(""))
            {
                if (MessageBox.Show("Continue SCR without Resolved Sheet", "Resolved Sheet not selected: SCR", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                {
                    path = textBox1.Text;
                    query = "select [SCR Details],[Product Categorization Tier1],[Product Categorization Tier2],[Incident: Number],[Resolved Date] from [<placeHolder>] where [Resolution Category Tier3] = 'New SCR' AND [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                    if (!dateTimePicker1.Enabled)
                    {
                        query = "select [SCR Details],[Product Categorization Tier1],[Product Categorization Tier2],[Incident: Number],[Resolved Date] from [<placeHolder>] where [Resolution Category Tier3] = 'New SCR'";// AND [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                    }
                }
                else
                {
                    throw new FileNotFoundException();
                }
            }
            else
            {
                path = textBox2.Text;
                query = "select [SCR Details],[Product Categorization Tier1],[Product Categorization Tier2],[Incident: Number],[Resolved Date] from [<placeHolder>] where [Resolution Category Tier3] = 'New SCR' AND [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                if (!dateTimePicker1.Enabled)
                {
                    query = "select [SCR Details],[Product Categorization Tier1],[Product Categorization Tier2],[Incident: Number],[Resolved Date] from [<placeHolder>] where [Resolution Category Tier3] = 'New SCR'";// AND [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                }
            }
            //Console.WriteLine("Entered Path: " + path);            
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;';"/*con*/))
            {
                connection.Open();
                OleDbCommand command;
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                // Change to remove Existing SCRs
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection);//command = new OleDbCommand("select [SCR Details],[Product Categorization Tier1],[Product Categorization Tier2],[Incident: Number],[Resolved Date] from [" + dtschema.Rows[0][2] + "] where [Resolution Category Tier3] = 'New SCR'", connection);
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //command = new OleDbCommand("select [SCR Details],[Product Categorization Tier1],[Product Categorization Tier2],[Incident: Number],[Resolved Date] from [" + dtschema.Rows[2][2] + "] where [Resolution Category Tier3] = 'New SCR'", connection);
                }
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);                
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    int dup = 0;
                    if (!dr["SCR Details"].Equals(DBNull.Value))
                        if (!(Convert.ToString(dr[0]).Trim().Equals("SCR Details")) && !(Convert.ToString(dr[0]).Trim().Equals("")))
                        {
                            foreach (DataRow drf in dtf.Rows)     // FOR REMOVING DUPES IN SCR
                            {
                                try
                                {
                                    if (drf["SCR Details"].Equals(dr["SCR Details"]))// || (drf["SCR Details"].ToString().Substring(0, drf["SCR Details"].ToString().Length - 2)).Equals(dr["SCR Details"].ToString().Substring(0, dr["SCR Details"].ToString().Length - 2)))
                                        dup++;
                                }
                                catch(ArgumentOutOfRangeException) 
                                {
                                    dup++;
                                }
                            }
                            if (dup == 0)
                                dtf.ImportRow(dr);
                        }
                }
                dv_temp = new DataView(dtf);
                dv_temp.Sort = "[Product Categorization Tier1],[Product Categorization Tier2] desc";
                //dv_temp.Table.TableName = "SCR Report";
            }
            //label11.Text = "\u2714";
            label11.Invoke((MethodInvoker)(() => label11.Text = "\u2714"));
            return dv_temp;
        }
        public DataView ByCustomer()
        {
            string query = "select * from [<placeHolder>] where [Opened Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            string query1 = "select * from [<placeHolder>] where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            if (!dateTimePicker1.Enabled)
            {
                query = "select * from [<placeHolder>]";// where [Opened Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
                query1 = "select * from [<placeHolder>]";// where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            }
            label10.Invoke((MethodInvoker)(() => label10.Text = "..."));
            string path1;
            DataView dv_temp;
            OleDbDataAdapter adaptor;
            DataTable dtf = new DataTable("By Customer");
            dtf.Columns.Add("Carrier");
            dtf.Columns.Add("P1");
            dtf.Columns.Add("P2");
            dtf.Columns.Add("P3");
            dtf.Columns.Add("P4");
            dtf.Columns.Add("P5");
            dtf.Columns.Add("Blank");
            dtf.Columns.Add("Total Issues");
            dtf.Columns.Add("P1R");
            dtf.Columns.Add("P2R");
            dtf.Columns.Add("P3R");
            dtf.Columns.Add("P4R");
            dtf.Columns.Add("P5R");
            dtf.Columns.Add("BlankR");
            dtf.Columns.Add("Issues Resolved");
            dtf.Columns.Add("% Resolved");
            DataTable dtf1 = new DataTable("By Customer");
            dtf1.Columns.Add("Carrier");
            dtf1.Columns.Add("P1");
            dtf1.Columns.Add("P2");
            dtf1.Columns.Add("P3");
            dtf1.Columns.Add("P4");
            dtf1.Columns.Add("P5");
            dtf1.Columns.Add("Blank");
            dtf1.Columns.Add("Total Issues");
            dtf1.Columns.Add("P1R");
            dtf1.Columns.Add("P2R");
            dtf1.Columns.Add("P3R");
            dtf1.Columns.Add("P4R");
            dtf1.Columns.Add("P5R");
            dtf1.Columns.Add("BlankR");
            dtf1.Columns.Add("Issues Resolved");
            dtf1.Columns.Add("% Resolved");
            path = textBox1.Text;
            if (textBox2.Text.Equals(""))
            {
                if (MessageBox.Show("Continue Customer without Resolved Sheet", "Resolved Sheet not selected: Customer", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                {
                    path1 = textBox1.Text;
                }
                else
                {
                    path1 = "";
                    throw new FileNotFoundException();
                }
            }
            else
            {
                path1 = textBox2.Text;
            }
            int found = -1;
            ArrayList category1 = new ArrayList();
            //Console.WriteLine("Entered Path: " + path);
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                DataSet ds = new DataSet();
                connection.Open();
                OleDbCommand command;
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                //OleDbCommand command = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "]", connection);
                //OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(dr["Customer"]))
                            found++;
                    }
                    if (found == -1)
                    {
                        category1.Add(dr["Customer"]);
                    }
                    else
                    {
                        found = -1;
                    }

                }
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                foreach (object cat in category1)
                {
                    if (cat.Equals(DBNull.Value))
                    {
                        dra = ds.Tables[0].Select("[Customer] IS NULL");
                    }
                    else
                        dra = ds.Tables[0].Select("[Customer]='" + cat + "'");


                    int p1 = 0, p2 = 0, p3 = 0, p4 = 0, p5 = 0, pb = 0;
                    int pr1 = 0, pr2 = 0, pr3 = 0, pr4 = 0, pr5 = 0, prb = 0;
                    foreach (DataRow dr in dra)
                    {
                        if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                        {
                            pb++;
                            if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                            {
                                prb++;
                            }
                        }
                        else
                        {
                            if (Convert.ToInt32(dr["Priority"]).Equals(1))
                            {
                                p1++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr1++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(2))
                            {
                                p2++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr2++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(3))
                            {
                                p3++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr3++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(4))
                            {
                                p4++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr4++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(5))
                            {
                                p5++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr5++;
                                }
                            }
                        }
                    }

                    DataRow drf = dtf.NewRow();

                    drf[0] = cat;//dra[0]["Customer"];
                    drf[1] = p1;
                    drf[2] = p2;
                    drf[3] = p3;
                    drf[4] = p4;
                    drf[5] = p5;
                    drf[6] = pb;
                    drf[7] = p1 + p2 + p3 + p4 + p5 + pb;
                    drf[8] = 0;// pr1;
                    drf[9] = 0;//pr2;
                    drf[10] = 0;//pr3;
                    drf[11] = 0;//pr4;
                    drf[12] = 0;//pr5;
                    drf[13] = 0;// prb;
                    drf[14] = 0;//pr1 + pr2 + pr3 + pr4 + pr5 + prb;
                    p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = 0;
                    dtf.Rows.Add(drf);
                }
            }
            found = -1;
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                DataSet ds = new DataSet();
                connection.Open();
                OleDbCommand command;
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query1.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query1.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                //OleDbCommand command = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "]", connection);
                //OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(dr["Customer"]))
                            found++;
                    }
                    if (found == -1)
                    {
                        category1.Add(dr["Customer"]);
                    }
                    else
                    {
                        found = -1;
                    }

                }
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                foreach (object cat in category1)
                {
                    if (cat.Equals(DBNull.Value))
                    {
                        dra = ds.Tables[0].Select("[Customer] IS NULL");
                    }
                    else
                        dra = ds.Tables[0].Select("[Customer]='" + cat + "'");


                    int p1 = 0, p2 = 0, p3 = 0, p4 = 0, p5 = 0, pb = 0;
                    int pr1 = 0, pr2 = 0, pr3 = 0, pr4 = 0, pr5 = 0, prb = 0;
                    foreach (DataRow dr in dra)
                    {
                        if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                        {
                            pb++;
                            if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                            {
                                prb++;
                            }
                        }
                        else
                        {
                            if (Convert.ToInt32(dr["Priority"]).Equals(1))
                            {
                                p1++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr1++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(2))
                            {
                                p2++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr2++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(3))
                            {
                                p3++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr3++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(4))
                            {
                                p4++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr4++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(5))
                            {
                                p5++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr5++;
                                }
                            }
                        }
                    }

                    DataRow drf = dtf1.NewRow();

                    drf[0] = cat;
                    drf[1] = 0;// p1;
                    drf[2] = 0;// p2;
                    drf[3] = 0;// p3;
                    drf[4] = 0;// p4;
                    drf[5] = 0;// p5;
                    drf[6] = 0;// pb;
                    drf[7] = 0;// p1 + p2 + p3 + p4 + p5 + pb;
                    drf[8] = pr1;
                    drf[9] = pr2;
                    drf[10] = pr3;
                    drf[11] = pr4;
                    drf[12] = pr5;
                    drf[13] = prb;
                    drf[14] = pr1 + pr2 + pr3 + pr4 + pr5 + prb;
                    p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = 0;
                    dtf1.Rows.Add(drf);
                }
            }
            List<string> listCarrier = new List<string>();
            foreach (DataRow drf in dtf1.Rows)
            {
                foreach (DataRow drT in dtf.Rows)
                {
                    if (drf["Carrier"].Equals(drT["Carrier"]))
                    {
                        drf["P1"] = drT["P1"];
                        drf["P2"] = drT["P2"];
                        drf["P3"] = drT["P3"];
                        drf["P4"] = drT["P4"];
                        drf["P5"] = drT["P5"];
                        drf["Blank"] = drT["Blank"];
                        drf["Total Issues"] = Convert.ToInt32(drT["P1"]) + Convert.ToInt32(drT["P2"]) + Convert.ToInt32(drT["P3"]) + Convert.ToInt32(drT["P4"]) + Convert.ToInt32(drT["P5"]) + Convert.ToInt32(drT["Blank"]);
                        listCarrier.Add(drT["Carrier"].ToString());
                    }
                }
            }
            foreach (DataRow drT in dtf.Rows)
            {
                if (!listCarrier.Contains(drT["Carrier"].ToString())/* && !catrgory_1.Contains(drT["Product Categorization Tier1"].ToString())*/)
                {
                    dtf1.ImportRow(drT);
                }
            }
            //label10.Text = "\u2714";
            label10.Invoke((MethodInvoker)(() => label10.Text = "\u2714"));
            dv_temp = new DataView(dtf1);
            dv_temp.Sort = "Carrier desc";
            return dv_temp;
        }
        public DataView EnvMetrics()
        {
            string query = "select * from [<placeHolder>] where [Opened Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            if (!dateTimePicker1.Enabled)
            {
                query = "select * from [<placeHolder>] where [Opened Date]";// between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            }
            label15.Invoke((MethodInvoker)(() => label15.Text = "..."));
            DataView dv_temp;
            OleDbDataAdapter adaptor;
            DataTable dtf = new DataTable("Environment Metrics");
            dtf.Columns.Add("Environment");            
            dtf.Columns.Add("Total Issues");
            DataTable dtf2 = new DataTable("Environment Metrics");
            dtf2.Columns.Add("Environment");
            dtf2.Columns.Add("Total Issues");
            path = textBox1.Text;
            int found = -1;
            ArrayList category1 = new ArrayList();
            //Console.WriteLine("Entered Path: " + path);
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                DataSet ds = new DataSet();
                connection.Open();
                OleDbCommand command;
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                //OleDbCommand command = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "]", connection);
                //OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(dr["Environment"]))
                            found++;
                    }
                    if (found == -1)
                    {
                        category1.Add(dr["Environment"]);
                    }
                    else
                    {
                        found = -1;
                    }

                }
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                foreach (object cat in category1)
                {
                    if (cat.Equals(DBNull.Value))
                    {
                        dra = ds.Tables[0].Select("[Environment] IS NULL");
                    }
                    else
                        dra = ds.Tables[0].Select("[Environment]='" + cat + "'");


                    int p1 = 0, p2 = 0, p3 = 0, p4 = 0, p5 = 0, pb = 0;
                    int pr1 = 0, pr2 = 0, pr3 = 0, pr4 = 0, pr5 = 0, prb = 0;
                    foreach (DataRow dr in dra)
                    {
                        if (dr["Priority"].Equals(DBNull.Value) || dr["Priority"].Equals("") || dr["Priority"].Equals("-"))
                        {
                            pb++;
                            if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                            {
                                prb++;
                            }
                        }
                        else
                        {
                            if (Convert.ToInt32(dr["Priority"]).Equals(1))
                            {
                                p1++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr1++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(2))
                            {
                                p2++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr2++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(3))
                            {
                                p3++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr3++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(4))
                            {
                                p4++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr4++;
                                }
                            }
                            if (Convert.ToInt32(dr["Priority"]).Equals(5))
                            {
                                p5++;
                                if (dr["Incident Status"].Equals("RESOLVED") || dr["Incident Status"].Equals("CLOSED"))
                                {
                                    pr5++;
                                }
                            }
                        }
                    }

                    DataRow drf = dtf.NewRow();
                    drf[0] = dra[0]["Environment"];                    
                    drf[1] = p1 + p2 + p3 + p4 + p5 + pb;                    
                    p1 = p2 = p3 = p4 = p5 = pb = pr1 = pr2 = pr3 = pr4 = pr5 = prb = 0;
                    dtf.Rows.Add(drf);
                }
            }
            int prod = 0, uat = 0;
            foreach (DataRow drf in dtf.Rows)
            {
                if (drf["Environment"].ToString().ToUpper().Contains("PROD"))
                {
                    prod += Convert.ToInt32(drf["Total Issues"]);
                }
                else if (drf["Environment"].ToString().ToUpper().Contains("UAT"))
                {
                    uat += Convert.ToInt32(drf["Total Issues"]);
                }
                else
                {
                    dtf2.ImportRow(drf);
                }
            }
            DataRow drt1 = dtf2.NewRow();
            drt1["Environment"] = "UAT";
            drt1["Total Issues"] = uat;
            dtf2.Rows.Add(drt1);
            DataRow drt = dtf2.NewRow();
            drt["Environment"] = "PROD";
            drt["Total Issues"] = prod;
            dtf2.Rows.Add(drt);
            dv_temp = new DataView(dtf2);
            dv_temp.Sort = "Environment desc";
            //label15.Text = "\u2714";
            label15.Invoke((MethodInvoker)(() => label15.Text = "\u2714"));
            return dv_temp;
        }
        public DataView IncidentStatus()
        {
            string query = "select * from [<placeHolder>] where [Opened Date] < #" + dateTimePicker2.Value + "#";
            if (!dateTimePicker1.Enabled)
            {
                query = "select * from [<placeHolder>]";// where [Opened Date] < #" + dateTimePicker2.Value + "#";
            }
            label16.Invoke((MethodInvoker)(() => label16.Text = "..."));
            OleDbDataAdapter adaptor;
            DataTable dtf = new DataTable("By Incident Status");
            dtf.Columns.Add("Product Categorization Tier1");
            dtf.Columns.Add("Product Categorization Tier2");
            dtf.Columns.Add("ASSIGNED");
            //dtf.Columns.Add("CANCELLED");
            dtf.Columns.Add("IN PROGRESS");
            dtf.Columns.Add("OPENED");
            dtf.Columns.Add("PENDING");
            //dtf.Columns.Add("RESOLVED");
            dtf.Columns.Add("WAITING FOR RESPONSE");
            dtf.Columns.Add("RESOLVED");
            dtf.Columns.Add("Blank");
            dtf.Columns.Add("Total Issues");
            if (File.Exists(textBox3.Text))
                path = textBox3.Text;
            else
            {
                path = textBox1.Text;
                MessageBox.Show("File 3 Does not Exist", "PresentlyOpened Sheet not selected: Going by Opened Sheet Instead");
            }
            int found = -1;
            ArrayList category1 = new ArrayList();
            ArrayList category2 = new ArrayList();
            //Console.WriteLine("Entered Path: " + path);
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                DataSet ds = new DataSet();
                DataSet ds_sub = new DataSet();
                OleDbCommand command;
                connection.Open();
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    try
                    {
                        foreach (object cat in category1)
                        {
                            if (cat.Equals(dr["Product Categorization Tier1"]))
                                found++;
                        }
                        if (found == -1)
                        {
                            category1.Add(dr["Product Categorization Tier1"]);
                        }
                        else
                        {
                            found = -1;
                        }
                    }
                    catch (InvalidCastException)
                    {
                        if (dr["Product Categorization Tier1"] == null)
                            category1.Add(null);
                    }
                }
                DataTable dt = ds.Tables[0].Clone();
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                try
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(DBNull.Value))
                        {
                            dra = ds.Tables[0].Select("[Product Categorization Tier1] IS NULL");
                        }
                        else
                            dra = ds.Tables[0].Select("[Product Categorization Tier1]='" + cat + "'");
                        //dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + cat + "'");//.CopyTo(drc,0);
                        foreach (DataRow drt in dra)
                        {
                            dt.ImportRow(drt);
                            try
                            {
                                foreach (object cat1 in category2)
                                {
                                    if (cat1.Equals(drt["Product Categorization Tier2"]))
                                        found++;
                                }
                                if (found == -1)
                                {
                                    category2.Add(drt["Product Categorization Tier2"]);
                                }
                                else
                                {
                                    found = -1;
                                }
                            }
                            catch (InvalidCastException)
                            {
                                if (drt["Product Categorization Tier2"] == null)
                                    category2.Add(null);
                            }
                        }
                        DataRow[] dra1;
                        int ass = 0, ip = 0, op = 0, pen = 0, b = 0, wfr = 0, res = 0;
                        foreach (object cat2 in category2)
                        {
                            if (cat2.Equals(DBNull.Value))
                            {
                                dra1 = dt.Select("[Product Categorization Tier2] IS NULL");
                            }
                            else
                                dra1 = dt.Select("[Product Categorization Tier2]='" + cat2 + "'");

                            foreach (DataRow dr in dra1)
                            {
                                if (dr["Incident Status"].Equals(DBNull.Value) || dr["Incident Status"].Equals(""))
                                {
                                    b++;
                                }
                                else
                                {
                                    if (dr["Incident Status"].Equals("ASSIGNED"))
                                    {
                                        ass++;
                                    }
                                    else if (dr["Incident Status"].Equals("IN PROGRESS"))
                                    {
                                        ip++;

                                    }
                                    else if (dr["Incident Status"].Equals("OPENED"))
                                    {
                                        op++;
                                    }
                                    else if (dr["Incident Status"].Equals("PENDING"))
                                    {
                                        pen++;
                                    }
                                    else if (dr["Incident Status"].Equals("WAITING FOR RESPONSE"))
                                    {
                                        wfr++;
                                    }
                                    else if (dr["Incident Status"].Equals("RESOLVED"))
                                    {
                                        res++;
                                    }
                                }
                            }
                            DataRow drf = dtf.NewRow();
                            drf[0] = dra1[0]["Product Categorization Tier1"];
                            drf[1] = dra1[0]["Product Categorization Tier2"];
                            drf[2] = ass;
                            //drf[3] = can;
                            drf[3] = ip;
                            drf[4] = op;
                            drf[5] = pen;
                            //drf[7] = res;
                            drf[6] = wfr;                            
                            drf[7] = res;
                            drf[8] = b;
                            drf[9] = ass + ip + op + pen + wfr + b + res;
                            ass = ip = op = pen = b = wfr = res = 0;
                            dtf.Rows.Add(drf);
                        }
                        dt.Clear();
                        category2.Clear();
                    }
                }
                catch (InvalidCastException)
                {
                    dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + null + "'");
                }
            }
            //label16.Text = "\u2714";
            label16.Invoke((MethodInvoker)(() => label16.Text = "\u2714"));
            DataView dv_temp;
            dv_temp = new DataView(dtf);
            dv_temp.Sort = "[Product Categorization Tier1] desc";
            return dv_temp;
        }
        public DataView LeakageRate()
        {
            string query = "select * from [<placeHolder>] where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            if (!dateTimePicker1.Enabled)
            {
                query = "select * from [<placeHolder>]";// where [Resolved Date] between #" + dateTimePicker1.Value + "# AND #" + dateTimePicker2.Value + "#";
            }
            label17.Invoke((MethodInvoker)(() => label17.Text = "..."));
            OleDbDataAdapter adaptor;
            DataTable dtf = new DataTable("Leakage Rate");
            dtf.Columns.Add("Product Categorization Tier1");
            dtf.Columns.Add("Product Categorization Tier2");
            dtf.Columns.Add("Escalated to DEV");
            path = textBox2.Text;
            int found = -1;
            ArrayList category1 = new ArrayList();
            ArrayList category2 = new ArrayList();
            //Console.WriteLine("Entered Path: " + path);
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
            {
                DataSet ds = new DataSet();
                DataSet ds_sub = new DataSet();
                OleDbCommand command;
                connection.Open();
                DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                try
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[0][2].ToString()), connection); //"select * from [" + dtschema.Rows[0][2] + "]"
                }
                catch (IndexOutOfRangeException)
                {
                    command = new OleDbCommand(query.Replace("<placeHolder>", dtschema.Rows[2][2].ToString()), connection);  //select * from [" + dtschema.Rows[2][2] + "]
                }
                adaptor = new OleDbDataAdapter(command);
                adaptor.Fill(ds);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    try
                    {
                        foreach (object cat in category1)
                        {
                            if (cat.Equals(dr["Product Categorization Tier1"]))
                                found++;
                        }
                        if (found == -1)
                        {
                            category1.Add(dr["Product Categorization Tier1"]);
                        }
                        else
                        {
                            found = -1;
                        }
                    }
                    catch (InvalidCastException)
                    {
                        if (dr["Product Categorization Tier1"] == null)
                            category1.Add(null);
                    }
                }
                DataTable dt = ds.Tables[0].Clone();
                DataRow[] dra;// = new DataRow[5000];                        
                found = -1;
                try
                {
                    foreach (object cat in category1)
                    {
                        if (cat.Equals(DBNull.Value))
                        {
                            dra = ds.Tables[0].Select("[Product Categorization Tier1] IS NULL");
                        }
                        else
                            dra = ds.Tables[0].Select("[Product Categorization Tier1]='" + cat + "'");
                        //dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + cat + "'");//.CopyTo(drc,0);
                        foreach (DataRow drt in dra)
                        {
                            dt.ImportRow(drt);
                            try
                            {
                                foreach (object cat1 in category2)
                                {
                                    if (cat1.Equals(drt["Product Categorization Tier2"]))
                                        found++;
                                }
                                if (found == -1)
                                {
                                    category2.Add(drt["Product Categorization Tier2"]);
                                }
                                else
                                {
                                    found = -1;
                                }
                            }
                            catch (InvalidCastException)
                            {
                                if (drt["Product Categorization Tier2"] == null)
                                    category2.Add(null);
                            }
                        }
                        DataRow[] dra1;
                        int esc = 0, b = 0;
                        foreach (object cat2 in category2)
                        {
                            if (cat2.Equals(DBNull.Value))
                            {
                                dra1 = dt.Select("[Product Categorization Tier2] IS NULL");
                            }
                            else
                                dra1 = dt.Select("[Product Categorization Tier2]='" + cat2 + "'");

                            foreach (DataRow dr in dra1)
                            {
                                if (dr["Incident Status"].Equals(DBNull.Value) || dr["Incident Status"].Equals(""))
                                {
                                    b++;
                                }
                                else
                                {
                                    if (dr["Resolution Category Tier2"].Equals("Escalated to DEV"))
                                    {
                                        esc++;
                                    }
                                }
                            }
                            DataRow drf = dtf.NewRow();
                            drf[0] = dra1[0]["Product Categorization Tier1"];
                            drf[1] = dra1[0]["Product Categorization Tier2"];
                            drf[2] = esc;
                            //drf[10] = ass + can + ip + op + pen + res + wfr + b;
                            //ass = can = ip = op = pen = b = res = wfr = 0;                            
                            if (esc != 0)
                                dtf.Rows.Add(drf);
                            esc = b = 0;
                        }
                        dt.Clear();
                        category2.Clear();
                    }
                }
                catch (InvalidCastException)
                {
                    dra = ds.Tables[0].Select(ds.Tables[0].Columns[2].Caption + "='" + null + "'");
                }
            }
            //label17.Text = "\u2714";
            label17.Invoke((MethodInvoker)(() => label17.Text = "\u2714"));
            DataView dv_temp;
            dv_temp = new DataView(dtf);
            dv_temp.Sort = "[Product Categorization Tier1] desc";
            return dv_temp;
        }
        public void genCompactSheet(string pathOfMonthlyReport, string pathOfResolvedDataFile)
        {
            string storagePathOfOutput=pathOfMonthlyReport.Replace(".xls","_Compacted.xls");
            label23.Invoke((MethodInvoker)(() => label23.Text = "..."));
            new MonthlyReportSummaryGenerator().ReadAndGenerateSummary(pathOfMonthlyReport, pathOfResolvedDataFile, storagePathOfOutput);
            label23.Invoke((MethodInvoker)(() => label23.Text = "\u2714"));
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel files (*.xls, *.xlsx)|*.xls*|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Process.Start(textBox1.Text);
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string aboutmsg = "SOAr Ver2.0\nIt is a Custom Reporting tool Which creates reports\nby using excel reports of specified format. \n\nAuthor: Sambhav Patni\nCourtesy: Infogain India Pvt. Ltd.";
            MessageBox.Show(aboutmsg, "About Me");
        }

        private void howToUseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string helpmsg = "Help In Progress...";
            MessageBox.Show("HELP YOURSELF!!!\n\nSorry, No Help Right Now...", helpmsg);
        }

        private void x64ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string helpmsg = "Using on 64 Bit";
            //MessageBox.Show("Follow this link and install: AccessDatabaseEngine to resolve the issue.\nhttp://www.microsoft.com/en-us/download/confirmation.aspx?id=23734", helpmsg);
            MessageBox.Show("Click Help to install: AccessDatabaseEngine to resolve the issue.", helpmsg, MessageBoxButtons.OK, MessageBoxIcon.Information,
    MessageBoxDefaultButton.Button3, 0, "http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734", "keyword");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel files (*.xls, *.xlsx)|*.xls*|All files (*.*)|*.*";
            openFileDialog1.ShowDialog();
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "Excel files (*.xls, *.xlsx)|*.xls*|All files (*.*)|*.*";
            openFileDialog2.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog3.Filter = "Excel files (*.xls, *.xlsx)|*.xls*|All files (*.*)|*.*";
            openFileDialog3.ShowDialog();
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            textBox2.Text = openFileDialog2.FileName;
        }

        private void openFileDialog3_FileOk(object sender, CancelEventArgs e)
        {
            textBox3.Text = openFileDialog3.FileName;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            backgroundWorker2.RunWorkerAsync();
        }
            /*
        {
            string ISTeam = "";
            string[] Temp;
            int Files_Processed = 0;
            OleDbDataAdapter adaptor;
            List<string> files = new List<string>();
            if (File.Exists(textBox1.Text))
                files.Add(textBox1.Text);
            if (File.Exists(textBox2.Text))
                files.Add(textBox2.Text);
            if (File.Exists(textBox3.Text))
                files.Add(textBox3.Text);
            if (checkBox8.Checked)
            {
                if (File.Exists(textBox1.Text))
                {
                    ProcessStartInfo theProcess = new ProcessStartInfo(textBox1.Text);
                    theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                    Process.Start(theProcess);
                }
                if (File.Exists(textBox2.Text))
                {
                    ProcessStartInfo theProcess = new ProcessStartInfo(textBox2.Text);
                    theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                    Process.Start(theProcess);
                }
                if (File.Exists(textBox3.Text))
                {
                    ProcessStartInfo theProcess = new ProcessStartInfo(textBox3.Text);
                    theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                    Process.Start(theProcess);
                }
            }
            try
            {
                Temp = File.ReadAllLines("ISOpsTeam.txt");
                foreach (string temp in Temp)
                {
                    ISTeam += "'" + temp + "',";
                }
                ISTeam = ISTeam.Substring(0, ISTeam.Length - 1);

                foreach (string Fpath in files)
                    using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Fpath + ";Extended Properties='Excel 12.0;HDR=YES;';"))
                    {
                        DataSet ds = new DataSet();
                        connection.Open();
                        OleDbCommand command;
                        OleDbCommand command1;
                        DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        try
                        {
                            command = new OleDbCommand("select * from [" + dtschema.Rows[0][2] + "] WHERE [Client ID] IN (" + ISTeam + ")", connection);
                            command1 = new OleDbCommand("select * from [" + dtschema.Rows[0][2] + "] WHERE [Client ID] NOT IN (" + ISTeam + ")", connection);
                        }
                        catch (IndexOutOfRangeException)
                        {
                            command = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "] WHERE [Client ID] IN (" + ISTeam + ")", connection);
                            command1 = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "] WHERE [Client ID] NOT IN (" + ISTeam + ")", connection);
                        }
                        adaptor = new OleDbDataAdapter(command);
                        adaptor.Fill(ds);
                        CreateWorkbook(Path.GetDirectoryName(Fpath) + "\\ISOPS-" + Fpath.Substring(Fpath.LastIndexOf("\\") + 1), ds);
                        ds = new DataSet();
                        adaptor = new OleDbDataAdapter(command1);
                        adaptor.Fill(ds);
                        CreateWorkbook(Path.GetDirectoryName(Fpath) + "\\TAC-" + Fpath.Substring(Fpath.LastIndexOf("\\") + 1), ds);
                        Files_Processed++;
                    }
            }
            catch (OleDbException)
            {
                MessageBox.Show("Excel Not in Proper Format:\n Try Selecting \"Open\" CheckBox , \n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com", "Error Occured");
            }
            catch (FileNotFoundException ex)
            {
                if (ex.Message.Contains("ISOpsTeam"))
                    MessageBox.Show("Config File Does Not Exist...\n" + ex.Message);
                else
                    MessageBox.Show("File Does Not Exist...\n" + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something Went Wrong...\n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com\n\n" + ex.Message, "Error Occured");
            }
            MessageBox.Show((Files_Processed * 2) + " Files have been Generated...", "Checkout");
        }
*/

        public void vba(string path)
        {
            label22.Invoke((MethodInvoker)(() => label22.Text = "..."));
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            //FileStream temp = File.OpenRead(path);
            string workbookPath = path;//temp.Name;
            //temp.Close();
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath,
                0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);
            var newStandardModule = excelWorkbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            var codeModule = newStandardModule.CodeModule;
            var lineNum = codeModule.CountOfLines + 1;
            var codeText = File.ReadAllText("Refresh.vb");
            codeModule.InsertLines(lineNum, codeText);
            //excelWorkbook.Save();
            //Console.WriteLine("Refresh");
            //File.AppendAllText(path_pre + "Log.txt", "macro_StartProcessing\n");            
            var macro = string.Format("{0}!{1}.{2}", excelWorkbook.Name, newStandardModule.Name, "Refresh");
            excelApp.Run(macro);            
            excelApp.Visible = true;            
            excelWorkbook.DoNotPromptForConvert = true;
            excelWorkbook.Save();
            excelWorkbook.Close();
            excelApp.Quit();
            label22.Invoke((MethodInvoker)(() => label22.Text = "\u2714"));
        }

        private void Do_Work_1(object sender, DoWorkEventArgs e)
        {
            //Moving it down ahould be absolute path
            String ISOpsTeam_file_path = "", AOMTeam_file_path = "";
            try
            {
                ISOpsTeam_file_path = File.OpenRead("ISOpsTeam.txt").Name;
                AOMTeam_file_path = File.OpenRead("AOMTeam.txt").Name;

                string ISTeam = "";
                string[] ISOpsTeam_list;
                string AOMTeam = "";
                string[] AOMTeam_list;
                string ALLTeam = "";

                button5.Invoke((MethodInvoker)(() => button5.Enabled = false));              

                int Files_Processed = 0;
                OleDbDataAdapter adaptor;
                List<string> files = new List<string>();
                if (File.Exists(textBox1.Text))
                    files.Add(textBox1.Text);
                if (File.Exists(textBox2.Text))
                    files.Add(textBox2.Text);
                if (File.Exists(textBox3.Text))
                    files.Add(textBox3.Text);
                if (checkBox8.Checked)
                {
                    if (File.Exists(textBox1.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox1.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                    if (File.Exists(textBox2.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox2.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                    if (File.Exists(textBox3.Text))
                    {
                        ProcessStartInfo theProcess = new ProcessStartInfo(textBox3.Text);
                        theProcess.WindowStyle = ProcessWindowStyle.Minimized;
                        Process.Start(theProcess);
                    }
                }
                label19.Invoke((MethodInvoker)(() => label19.Show()));
                try
                {
                    ISOpsTeam_list = File.ReadAllLines(ISOpsTeam_file_path);
                    foreach (string temp in ISOpsTeam_list)
                    {
                        ISTeam += "'" + temp + "',";
                    }                    

                    AOMTeam_list = File.ReadAllLines(AOMTeam_file_path);
                    foreach (string temp in AOMTeam_list)
                    {
                        AOMTeam += "'" + temp + "',";
                    }
                    
                        if (ISTeam.Length > 0)
                        {
                            ISTeam = ISTeam.Substring(0, ISTeam.Length - 1);
                            ALLTeam = ISTeam;
                        }
                        if (AOMTeam.Length > 0)
                        {
                            AOMTeam = AOMTeam.Substring(0, AOMTeam.Length - 1);
                            if (ISTeam.Length > 0)
                                ALLTeam += "," + AOMTeam;
                        }
                    

                    //ALLTeam = ISTeam + AOMTeam;

                    foreach (string Fpath in files)
                        using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Fpath + ";Extended Properties='Excel 12.0;HDR=YES;';"/*con*/))
                        {
                            DataSet ds = new DataSet();
                            connection.Open();
                            OleDbCommand command_IN_ISTeam;
                            OleDbCommand command_IN_AOMTeam;
                            OleDbCommand command_ExceptTeam;
                            DataTable dtschema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            try
                            {
                                if (ISTeam == "")
                                    ISTeam = "'--None--'";
                                if (AOMTeam == "")
                                    AOMTeam = "'--None--'";
                                if (ALLTeam == "")
                                    ALLTeam = "'--None--'";
                                command_IN_ISTeam = new OleDbCommand("select * from [" + dtschema.Rows[0][2] + "] WHERE [Client ID] IN (" + ISTeam + ")", connection);
                                command_IN_AOMTeam = new OleDbCommand("select * from [" + dtschema.Rows[0][2] + "] WHERE [Client ID] IN (" + AOMTeam + ")", connection);
                                command_ExceptTeam = new OleDbCommand("select * from [" + dtschema.Rows[0][2] + "] WHERE [Client ID] NOT IN (" + ALLTeam + ")", connection);
                            }
                            catch (IndexOutOfRangeException)
                            {
                                command_IN_ISTeam = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "] WHERE [Client ID] IN (" + ISTeam + ")", connection);
                                command_IN_AOMTeam = new OleDbCommand("select * from [" + dtschema.Rows[0][2] + "] WHERE [Client ID] IN (" + AOMTeam + ")", connection);
                                command_ExceptTeam = new OleDbCommand("select * from [" + dtschema.Rows[2][2] + "] WHERE [Client ID] NOT IN (" + ALLTeam + ")", connection);
                            }
                            adaptor = new OleDbDataAdapter(command_IN_ISTeam);
                            adaptor.Fill(ds);
                            string tFileName = Fpath.Substring(Fpath.LastIndexOf("\\") + 1);
                            tFileName = tFileName.Substring(0, tFileName.LastIndexOf('.') + 1);
                            tFileName += "xls";
                            string tPath = Path.GetDirectoryName(Fpath) + "\\ISOPS-" + tFileName;
                            CreateWorkbook(tPath, ds);

                            ds = new DataSet();
                            adaptor = new OleDbDataAdapter(command_IN_AOMTeam);
                            adaptor.Fill(ds);
                            tPath = Path.GetDirectoryName(Fpath) + "\\AOM-" + tFileName;
                            CreateWorkbook(tPath, ds);

                            ds = new DataSet();
                            adaptor = new OleDbDataAdapter(command_ExceptTeam);
                            adaptor.Fill(ds);
                            tPath = Path.GetDirectoryName(Fpath) + "\\TAC-" + tFileName;
                            CreateWorkbook(tPath, ds);
                            Files_Processed++;
                            if (Files_Processed == 1)
                            {
                                textBox1.Invoke((MethodInvoker)(() => textBox1.Text = tPath));
                                checkBox12.Invoke((MethodInvoker)(() => checkBox12.Checked = false));
                            }
                            if (Files_Processed == 2)
                            {
                                textBox2.Invoke((MethodInvoker)(() => textBox2.Text = tPath));
                                checkBox12.Invoke((MethodInvoker)(() => checkBox12.Checked = false));
                            }
                            if (Files_Processed == 3)
                            {
                                textBox3.Invoke((MethodInvoker)(() => textBox3.Text = tPath));
                                checkBox12.Invoke((MethodInvoker)(() => checkBox12.Checked = false));
                            }
                        }
                }
                catch (OleDbException)
                {
                    MessageBox.Show("Excel Not in Proper Format:\n Try Selecting \"Open\" CheckBox , \n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com", "Error Occured");
                }
                catch (FileNotFoundException ex)
                {
                    if (ex.Message.Contains("ISOpsTeam"))
                        MessageBox.Show("Config File Does Not Exist...\n" + ex.Message);
                    else
                        MessageBox.Show("File Does Not Exist...\n" + ex.Message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Something Went Wrong...\n\n Contact: Sambhav Patni\n At: Sambhav.patni@infogain.com\n\n" + ex.Message, "Error Occured");
                }
                label19.Invoke((MethodInvoker)(() => label19.Hide()));
                MessageBox.Show((Files_Processed * 3) + " Files have been Generated...", "Checkout");
                button5.Invoke((MethodInvoker)(() => button5.Enabled = true));
            }
            catch { MessageBox.Show("ISOpsTeam/AOMTeam text file missing.", "Files missing", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox4.Text = dateTimePicker1.Value.Day + Month[dateTimePicker1.Value.Month-1].Substring(0, 3) + " - " + dateTimePicker2.Value.Day + Month[dateTimePicker2.Value.Month-1].Substring(0, 3); // 26/5 - 25/12
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            textBox4.Text = dateTimePicker1.Value.Day + Month[dateTimePicker1.Value.Month-1].Substring(0, 3) + " - " + dateTimePicker2.Value.Day + Month[dateTimePicker2.Value.Month-1].Substring(0, 3); // 26/5 - 25/12
        }

        private void checkBox12_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox12.Checked)
            {
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
            }
            else
            {
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }
    }
}
