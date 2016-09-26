using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class PersistentExcelRange
    {
        public ExcelRange Range
        {
            get
            {
                return sheet.Cells[startRow, startCol, endRow, endCol];
            }
        }
        public ExcelCellAddress Start { get { return Range.Start; } }
        public ExcelCellAddress End { get { return Range.End; } }

        private ExcelWorksheet sheet;
        private int startRow;
        private int startCol;
        private int endRow;
        private int endCol;

        public PersistentExcelRange()
        {
        }

        public PersistentExcelRange(ExcelWorksheet sheet)
        {
            this.setRange(sheet,
                            sheet.Dimension.Start.Row,
                            sheet.Dimension.Start.Column,
                            sheet.Dimension.End.Row,
                            sheet.Dimension.End.Column
            );

        }
        public PersistentExcelRange(ExcelWorksheet sheet, int Row)
        {
            this.setRange(sheet,
                            Row,
                            sheet.Dimension.Start.Column,
                            Row,
                            sheet.Dimension.End.Column
            );
        }

        public PersistentExcelRange(ExcelWorksheet sheet, int startRow, int endRow)
        {
            this.setRange(sheet,
                            startRow,
                            sheet.Dimension.Start.Column,
                            endRow,
                            sheet.Dimension.End.Column
            );
        }

        public PersistentExcelRange(ExcelWorksheet sheet, int startRow, int startCol, int endRow, int endCol)
        {
            this.setRange(sheet, startRow, startCol, endRow, endCol);
        }

       

        public void setRange(ExcelWorksheet sheet, int startRow, int startCol, int endRow, int endCol)
        {
            this.sheet = sheet;
            this.startRow = startRow;
            this.startCol = startCol;
            this.endRow = endRow;
            this.endCol = endCol;
        }
        /// <summary>
        /// Returns if the specified cell contains an excel formula
        /// </summary>
        /// <param name="row">Row index of cell</param>
        /// <param name="col">Column Index of Cell</param>
        /// <param name="offset">If true, then previous values are considered as offset to starting row and starting column index</param>
        /// <returns>True, if cell has a formula. Otherwise false</returns>
        public bool IsFormula(int row, int col, bool offset=false)
        {
            if (offset) {
                OffsetIndices(ref row, ref col);
            }            
            return this.Range[row, col].IsFormula();
        }

        private void OffsetIndices(ref int row, ref int col)
        {
           
            OffsetColumnIndex(ref col);
            OffsetRowIndex(ref row);
        }
        
        public bool IsFormula(int col, bool offset=false)
        {
            if (offset) {
                OffsetColumnIndex(ref col);            
            }
                
            return this.IsFormula(this.Start.Row, col);
        }

        private void OffsetColumnIndex(ref int col)
        {
            col += this.Start.Column;            
        }
        private void OffsetRowIndex(ref int row)
        {            
            row += this.Start.Row;            
        }

        /// <summary>
        /// Gives Excel cell reference
        /// </summary>
        /// <param name="row">Row index of cell</param>
        /// <param name="col">Column Index of Cell</param>
        /// <param name="offset">If true, then previous values are considered as offset to starting row and starting column index</param>
        /// <returns>ExcelRange reference to Cell</returns>
        public ExcelRange this[int row, int col, bool offset=false]
        {
            get
            {
                if (offset) {
                    OffsetIndices(ref row, ref col);
                }              
                return this.Range[row, col];
            }
        }

        /// <summary>
        /// Returns Excel Cell object at specified column in first row
        /// </summary>
        /// <param name="col"> Column Index</param>
        /// <param name="offset">if true, then column index is considered to be offset to starting column index</param>
        /// <returns></returns>
        public ExcelRange this[int col, bool offset = false]
        {
            get
            {                
                if (offset) {
                    OffsetColumnIndex(ref col);
                }
                return this.Range[ this.Start.Row, col];
            }
        }
        /// <summary>
        /// returns cell range corresponding to row and column indexes 
        /// </summary>
        /// <param name="s_row"> Row Start index.(inclusive)</param>
        /// <param name="s_col">Column Start Index (inclusive)</param>
        /// <param name="e_row"> Row End index (inclusive)</param>
        /// <param name="e_col">Column End Index(inclusive)</param>
        /// <param name="offset">Default false. If it is true then previous index are treated as offset to this.Start.Row and this.Start.Column</param>
        /// <returns>ExcelRange Corresponding to selection</returns>
        public ExcelRange this[int s_row, int s_col, int e_row, int e_col, bool offset=false]
        {
            get
            {
                if (offset) {
                    OffsetIndices(ref s_row, ref s_col);
                    OffsetIndices(ref e_row, ref e_col);
                }             
                
                return this.Range[s_row, s_col, e_row, e_col];
            }
        }

    }


    public class ResolvedReportRowData
    {
        public string PrimaryTier { get; set; }
        public string SecondaryTier { get; set; }
        public int Priority { get; set; }
        public TimeSpan timeTaken { get; set; }
        public int Count { get; set; }

    }
}
