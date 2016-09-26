using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Reflection;

namespace ExcelBeautifier
{
    public static class ExcelWorkSheetExtensions
    {
       public static bool IsFormula(this ExcelRange range)
        {
            return range.Formula != "";
        }
       public static void CopyPropertiesTo<T>(this T source, T dest)
       {
           var plist = from prop in typeof(T).GetProperties() where prop.CanRead && prop.CanWrite select prop;

           foreach (PropertyInfo prop in plist) {
               prop.SetValue(dest, prop.GetValue(source, null), null);
           }
       }
     
    }

   
}
