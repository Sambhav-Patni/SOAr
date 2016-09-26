using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    public class ReportRow
    {
        public string Tier1Name { get; set; }
        public string Tier2Name { get; set; }
        public PersistentExcelRange DataRange { get; set; }



        public int RowsContained { get; set; }

        public ReportRow()
        {
            RowsContained = 1;
        }

        public bool IsFormula(int row, int col, bool offset=false)
        {
            return this.DataRange.IsFormula(row, col,offset);
        }
        public bool IsFormula(int col, bool offset=false)
        {
           
            return this.DataRange.IsFormula(col, offset);
        }

        public bool isSameAs(ReportRow that)
        {
            if (this == that)
                return true;

            if (
                this.Tier1Name == that.Tier1Name &&
                this.Tier2Name == that.Tier2Name &&
                this.DataRange.Start.Row == that.DataRange.Start.Row
            ) {
                return true;
            }

            return false;

        }

    }
}
