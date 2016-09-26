using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class ExcelSheetDto
    {
        private NPOI.SS.UserModel.ISheet sheetToSummarize;

        public string Name { get; set; }
        

        public ExcelSheetDto()
        {
            this.Headers = new List<ExcelRowDto>();
            this.Footer = new ExcelRowDto();
            this.Body = new List<ExcelRowDto>();
        }
      
        public void UpdateFormula(ExcelCellDto ecell)
        {
           
        }
        
        public List<ExcelRowDto> Headers { get; set; }
        public List<ExcelRowDto> Body { get; set; }
        public ExcelRowDto Footer { get; set; }



        internal void UpdateAllForumlas()
        {
            //update formulas on body
            for (int row_it = 0; row_it < this.Body.Count; row_it++) {
                foreach (var cell in this.Body[row_it].ListOfCells) {
                    if (cell.Type == CellType.Formula) {
                        
                        var formula = cell.val as CellFormula;
                        if (formula.isSum == true) {
                            throw new InvalidOperationException("Sum formula should not be in body");    
                        }
                        formula.updateCeilingFormula( row_it + this.Headers.Count);
                    }
                }    
            }

            //updates formulas on footer if footer exists
            if (this.Footer == null)
                return;

            foreach (var cell in this.Footer.ListOfCells) {
                if (cell.Type == CellType.Formula) {
                    int row_num = this.Headers.Count + this.Body.Count;
                    var formula = cell.val as CellFormula;
                    if (formula.isSum == true) {
                        formula.updateSumFormula(this.Headers.Count+1, row_num);
                    }
                    else {
                        formula.updateCeilingFormula(row_num);
                    }
                }
            }    


        }
    }
}
