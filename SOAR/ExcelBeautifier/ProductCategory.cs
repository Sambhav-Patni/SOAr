using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class ProductCategory
    {
        public string PrimaryTierName { get; set; }
        public string SecondaryTierName { get; set; }
        public List<ExcelCellDto> Properties { get; set; }
        public ProductCategory()
        {
            Properties = new List<ExcelCellDto>();
        }
    }

}
