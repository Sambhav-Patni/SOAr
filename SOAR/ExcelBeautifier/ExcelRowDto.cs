using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class ExcelRowDto
    {
        public ExcelRowDto()
        {

            ListOfCells = new List<ExcelCellDto>();
        }
        public List<ExcelCellDto> ListOfCells { get; set; }       
    }
}
