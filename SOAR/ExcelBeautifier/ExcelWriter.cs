using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class ExcelWriter
    {

        internal static NPOI.SS.UserModel.IWorkbook ToExcelWorkbook(List<ExcelSheetDto> listOfSummarizedExcelData)
        {

            //create workbook
            var objToReturn = new HSSFWorkbook();
             //write header
                         
        

            //foreach sheet create an NPOI sheet and add it to workbook
            foreach (var sheetData in listOfSummarizedExcelData) {
                var sheetToReturn = objToReturn.CreateSheet(sheetData.Name);
                int rowNum = 0;

                sheetData.UpdateAllForumlas();
                foreach (var sheetRow in sheetData.Headers) {
                    WriteExcelRow(sheetToReturn, sheetRow, ref rowNum);
                }
                string prevTier1 = "";
                string prevTier2 = "";
                foreach (var sheetRow in sheetData.Body) {
                    if (sheetRow.ListOfCells[0].GetValueAsString() == prevTier1) {
                        sheetRow.ListOfCells[0].SetValue("");
                    }
                    else {
                        prevTier1 = sheetRow.ListOfCells[0].GetValueAsString();
                    }

                    if (sheetRow.ListOfCells[1].GetValueAsString() == prevTier2) {
                        sheetRow.ListOfCells[1].SetValue("");
                    }
                    else {
                        prevTier2 = sheetRow.ListOfCells[1].GetValueAsString();
                    }


                    WriteExcelRow(sheetToReturn, sheetRow, ref rowNum);
                    
                }

                if (sheetData.Footer != null) {
                    WriteExcelRow(sheetToReturn, sheetData.Footer, ref rowNum);
                }
               

            }

            //write footer
            

            return objToReturn;
        }

        private static void WriteExcelRow( NPOI.SS.UserModel.ISheet sheetToReturn, ExcelRowDto sheetRow, ref int rowNum)
        {
            var rowToReturn_T = sheetToReturn.CreateRow(rowNum);
            rowNum++;

            for (int ci_io = 0; ci_io < sheetRow.ListOfCells.Count; ci_io++) {
                var newCell = rowToReturn_T.CreateCell(ci_io);
                sheetRow.ListOfCells[ci_io].CopyTo(newCell);
            }
           
        }
    }

}
