using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelBeautifier
{

    public enum CellType
    {
        Unknown = -1,
        Numeric = 0,
        String = 1,
        Formula = 2,
        Blank = 3,
        Boolean = 4,
        Error = 5,
        DateTime = 6,
        TimeSpan = 7

    }
   
    public class CellFormula
    {
        public string Name { get; private set; }
        public int Row1 { get; set; }
        public int Row2 { get; set; }
        public string Col1 { get; set; }
        public string Col2 { get; set; }
        public bool isSum { get; set; }

        public int OriginRow { get; set; }
        public string OriginColumn { get; set; }
        public string FormulaText { get; set; }

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0) {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
      
        public string GetSumFormula(int row1, int row2)
        {
            string column =OriginColumn;
            return "SUM("+column+ row1+":"+column+ row2+")";
        }
        public string GetCielingFormula(int col1, int col2)
        {
            string column1 = GetExcelColumnName(col1);
            string column2 = GetExcelColumnName(col2);

            return GetCeilingFormula(column1, column2);
        }

        public string GetCeilingFormula(string column1, string column2)
        {
            return "CEILING( (" + column1 + OriginRow + " *100/ " + column2 + OriginRow + " ),1)";
        }


        public static CellFormula Create(string excelFormula, int orig_row, int orig_col)
        {
            CellFormula objToReturn_R = new CellFormula();
            objToReturn_R.OriginColumn = GetExcelColumnName(orig_col + 1); //+1 cause excel is 1-based
            objToReturn_R.OriginRow = orig_row + 1; 

            objToReturn_R.isSum = false;
            if (excelFormula.ToUpper().Contains("SUM")) {
                objToReturn_R.isSum = true;
            }
            Match match = Regex.Match(excelFormula,  @"([A-Z]{1,2}[0-9]{1,2})");
           
            //there must be 2 cell addresses in excelFormula
            string cell1addr = match.Value;
            string cell2addr = match.NextMatch().Value;

            int row_t;
            string col_t;
            SplitExcelAddress(cell1addr, out row_t, out col_t);
            objToReturn_R.Col1 = col_t;
            objToReturn_R.Row1 = row_t;

            SplitExcelAddress(cell2addr, out row_t, out col_t);
            objToReturn_R.Col2 = col_t;
            objToReturn_R.Row2 = row_t;


            //Console.WriteLine(match.Value);
            return objToReturn_R;
            
        }

        private static void SplitExcelAddress(string cell1addr, out int row_t, out string col_t)
        {
            row_t = Convert.ToInt32(Regex.Match(cell1addr, @"[0-9]{1,4}").Value);
            col_t = Regex.Match(cell1addr, @"[A-Z]{1,4}").Value;
        }


        internal void updateCeilingFormula(int row_it)
        {
            this.OriginRow = row_it + 1;
            this.FormulaText = this.GetCeilingFormula(this.Col1, this.Col2);
        }

        internal void updateSumFormula(int p1, int p2)
        {
            this.FormulaText = this.GetSumFormula(p1,p2);
        }
    }

    public class ExcelCellDto
    {
        public object val { get; private set; }
        public CellType Type { get; private set; }
        public string Name { get; set; }
        
        

        public ExcelCellDto()
        {
            this.Type = CellType.String;
            this.val = "";
            this.Name = "";
        }

        public ExcelCellDto(string val)
        {
            // TODO: Complete member initialization
            this.val = val;
            this.Type = CellType.String;
        }

        public ExcelCellDto(double val)
        {
            // TODO: Complete member initialization
            this.val = val;
            this.Type = CellType.Numeric;
        }

        public T GetValue<T>()
        {           
            return (T)val;
        }


        public string GetValueAsString()
        {
            string text = val.ToString();
            if (Type == CellType.Formula) {
                text = "=" + text;
            }                
            return text;
        }

        public TimeSpan GetValueAsTimeSpan()
        {
            var ticks = (long)val;
            if (Type == CellType.TimeSpan) {
                return new TimeSpan(ticks);
            }
            throw new InvalidOperationException("The cell type is not of timespan type");
        }

        public void SetValue(string val, bool isFormula=false)
        {
            this.val = val;
            this.Type = CellType.String;
            if (isFormula) {
                this.Type = CellType.Formula;
            }
        }

        public void SetValue(double val)
        {
            this.val = val;
            this.Type = CellType.Numeric;            
        }

        public void AddValue(double val)
        {
            
            if(this.Type == CellType.Numeric)
              this.val = (double)this.val + val;

            if (this.Type == CellType.String && this.val == "") {
                this.val = val;
                this.Type = CellType.Numeric;
            }
        }



        internal static ExcelCellDto CreateFrom(NPOI.SS.UserModel.ICell cell)
        {
            if(cell == null) 
                return new ExcelCellDto();

            var cellToReturn_R = new ExcelCellDto();
            
            cellToReturn_R.Type = ToCellType( cell.CellType);

            switch (cellToReturn_R.Type) {
                case CellType.Blank:
                    cellToReturn_R.val = "";
                    break;
                case CellType.Boolean:
                    cellToReturn_R.val = cell.BooleanCellValue;
                    break;
                case CellType.Error:
                    cellToReturn_R.val = cell.ErrorCellValue;
                    break;
                case CellType.Formula:
                   cellToReturn_R.Type = CellType.Formula;
                   cellToReturn_R.val = CellFormula.Create( cell.CellFormula, cell.RowIndex, cell.ColumnIndex);
                    break;
                case CellType.Numeric:
                    cellToReturn_R.val = cell.NumericCellValue;                    
                    break;
                case CellType.String:
                    cellToReturn_R.val = cell.StringCellValue;
                     break;
                case CellType.DateTime:
                     cellToReturn_R.val = cell.DateCellValue;
                     break;

                case CellType.TimeSpan:
                     cellToReturn_R.val = cell.DateCellValue.Ticks;
                     break;

                case CellType.Unknown:
                     cellToReturn_R.val = cell;
                    break;
                default:
                    break;
            }

            return cellToReturn_R;
        }

        private static CellType ToCellType(NPOI.SS.UserModel.CellType cellType)
        {
            CellType t = (CellType)((int)cellType);
            return t;
        }



        internal void CopyTo(ICell eCell)
        {
            eCell.SetCellType ( (NPOI.SS.UserModel.CellType)((int) this.Type));

            if (this.Type == CellType.DateTime || this.Type == CellType.TimeSpan) {
               
            }

            switch (this.Type) {
                case CellType.Unknown:
                    
                    break;
                case CellType.Numeric:
                    eCell.SetCellValue((double)this.val);
                    break;
                case CellType.String:

                    eCell.SetCellValue((string)this.val);
                    break;
                case CellType.Formula:
                    eCell.SetCellType(NPOI.SS.UserModel.CellType.Formula);
                    //throw new NotImplementedException();
                    var formula = (CellFormula)this.val;
                    eCell.SetCellFormula( formula.FormulaText);
                    
                    break;
                case CellType.Blank:
                    eCell.SetCellValue("");
                    break;
                case CellType.Boolean:
                    eCell.SetCellValue((bool) this.val);
                    break;
                case CellType.Error:
                    eCell.SetCellErrorValue((byte)this.val);
                    break;
                case CellType.DateTime:
                    eCell.SetCellValue((DateTime) this.val);
                    eCell.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                    IDataFormat dataFormatCustom = eCell.Row.Sheet.Workbook.CreateDataFormat();               
                    eCell.CellStyle.DataFormat = dataFormatCustom.GetFormat("yyyyMMdd HH:mm:ss");            
                    break;
                case CellType.TimeSpan:
                    var offdate = new DateTime(1900,1,1);
                    offdate.AddDays(-1); //because excel thinks 1900 was a leap year.
                    offdate.AddDays( ((TimeSpan)this.val).TotalDays);
                    eCell.SetCellValue(offdate);    //BUG: this may cause error. Yet to run the code
                     eCell.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                    IDataFormat timeFormatCustom = eCell.Row.Sheet.Workbook.CreateDataFormat();               
                    eCell.CellStyle.DataFormat = timeFormatCustom.GetFormat("DD:HH:mm:ss");            
                    
                    break;
                default:
                    break;
            }
        }
    }
}
