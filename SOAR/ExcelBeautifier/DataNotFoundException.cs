using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelBeautifier
{
    class DataNotFoundException : Exception
    {
        public DataNotFoundException(string Message ="", Exception Inner=null ): base(Message, Inner)
        {          

        }
    }
}
