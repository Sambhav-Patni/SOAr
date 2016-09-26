using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelBeautifier
{
    public static class ApdMonthlyReportHelper
    {
        /// <summary>
        ///  Transform a SheetToSummarize to Corresponding ExcelSheetDto 
        /// </summary>
        /// <param name="sheetToSummarize"></param>
        /// <returns></returns>
        internal static ExcelSheetDto SanitizeAndTransform(NPOI.SS.UserModel.ISheet sheetToSummarize)
        {
            NormalizeReportSheet(sheetToSummarize);

            var listOfRowsOfExcelCellsInSheet_Tp = sheetToSummarize.AsListOfExcelRowDto();
            var sheetDtoToReturn_Rn = listOfRowsOfExcelCellsInSheet_Tp.AsExcelSheetDto();

            return sheetDtoToReturn_Rn;
        }

        private static ExcelSheetDto AsExcelSheetDto(this List<ExcelRowDto> listOfRowsToConvert_Pi)
        {
            ExcelSheetDto objectToReturn_Rn = new ExcelSheetDto();

            //usually Excel shhet's header will consist of 2 rows
            int countOfRowsInHeader = 2;

            //but it may be that 2nd row contains data instead of heading or blank
            if (listOfRowsToConvert_Pi.Count > 1) {
                var cellToMatch = listOfRowsToConvert_Pi[1].ListOfCells[0];
                var textToMatch = ExcelReportSummaryConfiguration.NameOfHeadingOfPrimaryTierColumn;

                if (cellToMatch.IsEmptyOrContains(textToMatch) == false) {
                    countOfRowsInHeader = 1;
                }
            }

            ExcelRowDto FooterRow = null;
            var lastRowOfSheet = listOfRowsToConvert_Pi.Last();
            if (lastRowOfSheet.IsFooterRow()) {
                FooterRow = lastRowOfSheet;
                listOfRowsToConvert_Pi.Remove(lastRowOfSheet);
            }
            objectToReturn_Rn.Footer = FooterRow;
            objectToReturn_Rn.Headers.AddRange(listOfRowsToConvert_Pi.Take(countOfRowsInHeader));

            objectToReturn_Rn.Body.AddRange(listOfRowsToConvert_Pi.Skip(countOfRowsInHeader));
            //Sanitize body
            string prevTier1Name = "";
            string prevTier2Name = "";
            var body = objectToReturn_Rn.Body;

            for (int i_it = 0; i_it < body.Count; i_it++) {
                if (body[i_it].ListOfCells[0].GetValueAsString() == "") {
                    body[i_it].ListOfCells[0].SetValue(prevTier1Name);
                }
                if (body[i_it].ListOfCells[1].GetValueAsString() == "") {
                    body[i_it].ListOfCells[1].SetValue(prevTier2Name);
                }
                prevTier1Name = body[i_it].ListOfCells[0].GetValueAsString();
                prevTier2Name = body[i_it].ListOfCells[1].GetValueAsString();

            }


            return objectToReturn_Rn;
        }

        private static bool IsFooterRow(this ExcelRowDto rowOfExcelSheet_Pi)
        {
            return rowOfExcelSheet_Pi.ListOfCells[0].GetValueAsString() == "" &&
                           rowOfExcelSheet_Pi.ListOfCells[1].GetValueAsString() == "Total";
        }        

        private static bool IsEmptyOrContains(this ExcelCellDto cellToMatch, string textToMatch)
        {
            bool isSameOrBlank = false;

            if (cellToMatch.GetValueAsString() == textToMatch || cellToMatch.GetValueAsString() == ""
            ) {
                isSameOrBlank = true;
            }
            return isSameOrBlank;
        }


        private static List<ExcelRowDto> AsListOfExcelRowDto(this NPOI.SS.UserModel.ISheet sheetToSummarize)
        {
            var objectToReturn_R = new List<ExcelRowDto>();
            string previousTier1Name = "";
            string previousTier2Name = "";

            for (int rowIndex_Kz = 0; rowIndex_Kz <= sheetToSummarize.LastRowNum; rowIndex_Kz++) {

                var rowForInput = sheetToSummarize.GetRow(rowIndex_Kz);

                if (rowForInput == null) {
                    objectToReturn_R.Add(new ExcelRowDto());
                    continue;
                }

                ExcelRowDto rowToReturn_T = new ExcelRowDto();

                rowToReturn_T = AsExcelRowDto(rowForInput);
                //  if(rowToReturn_T.ListOfCells[0].GetValueAsString() == "")

                objectToReturn_R.Add(rowToReturn_T);
            }
            return objectToReturn_R;
        }

        private static ExcelRowDto AsExcelRowDto(NPOI.SS.UserModel.IRow rowForInput)
        {
            var objectToReturn_R = new ExcelRowDto();
            for (int celIndex_Kz = 0; celIndex_Kz < rowForInput.LastCellNum; celIndex_Kz++) {
                var cell = rowForInput.GetCell(celIndex_Kz);
                objectToReturn_R.ListOfCells.Add(ExcelCellDto.CreateFrom(cell));
            }
            return objectToReturn_R;
        }

        private static void NormalizeReportSheet(NPOI.SS.UserModel.ISheet sheetToSummarize)
        {
            if (sheetToSummarize.SheetName == ExcelReportSummaryConfiguration.NameOfScrReportSheet) {
                //delete first 2 columns because data in this sheet starts from column 3
                // index starts from 0 


                List<int> indicesOfCellsToDelete = new List<int>() { 0, 1 };
                for (int i_it =0; i_it  <= sheetToSummarize.LastRowNum ; i_it++) {
                    var row =   sheetToSummarize.GetRow(i_it);
                    //first 2 columns contain unwantwed data hence are removed
                    DeleteFirst2Column(row);

                    // for scr report, we are storing count of rows and overwriting actual value
                    //this is buisness logic
                    if (i_it != 0) {
                        row.GetCell(2).SetCellValue(1); 
                    }
               }
              

            }

        }

        //wasted 5+ hours because
        //I could not use Epplus
        //:(              
        private static void DeleteFirst2Column(NPOI.SS.UserModel.IRow row)
        {
            row.RemoveCell(row.GetCell(0));
            row.RemoveCell(row.GetCell(1));

            row.MoveCell(row.GetCell(2), 0);
            row.MoveCell(row.GetCell(3), 1);
            row.MoveCell(row.GetCell(4), 2);
        }



       


        public static List<NPOI.SS.UserModel.ICell> RemoveAt(List<NPOI.SS.UserModel.ICell> listOfCells_T, List<int> indicesOfCellsToDelete)
        {
            var objectToReturn_R = listOfCells_T.Where(
                                        cell => indicesOfCellsToDelete.Contains(cell.ColumnIndex) != true
                                    ).ToList();

            return objectToReturn_R;
        }
    }
}
