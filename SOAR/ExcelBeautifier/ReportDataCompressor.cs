using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBeautifier
{
    //to compress data provided by through ExcelDto 
    //according to rules of Configuration provided
   static public class ReportDataSummarizer
    {

     

        //interfacing the MonthlyReportSummarizer
      
        internal static List<ExcelRowDto> Summarize(List<ExcelRowDto> list_Pi)
        {
            List<ProductCategory> listOfProductCategories_T = list_Pi.Select(rowDto => rowDto.ToProductCategory()).ToList();

            listOfProductCategories_T = compressProductCategories(listOfProductCategories_T);

            List<ExcelRowDto> objectToReturn_R = listOfProductCategories_T.Select( pc => ToExcelRow(pc)).ToList();

            return objectToReturn_R;
        }

        private static ExcelRowDto ToExcelRow(ProductCategory pc)
        {
            ExcelRowDto objectToReturn_R = new ExcelRowDto();

            objectToReturn_R.ListOfCells.Add(new ExcelCellDto(pc.PrimaryTierName));
            objectToReturn_R.ListOfCells.Add(new ExcelCellDto(pc.SecondaryTierName));
            objectToReturn_R.ListOfCells.AddRange(pc.Properties);

            return objectToReturn_R;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="objectToCompress_Pi"></param>
        /// <returns></returns>
        private static List<ProductCategory> compressProductCategories(List<ProductCategory> source_Pi)
        {
           
            var storage_T = source_Pi.Select(pc => new ProductCategory() {
                PrimaryTierName = pc.FindMasterProductCategory().PrimaryTierName,
                SecondaryTierName = pc.FindMasterProductCategory().SecondaryTierName,
                Properties = pc.Properties            
            }).ToList();

            var groups_T =  storage_T.GroupBy(  
                                pc => new {
                                    tier1 = pc.PrimaryTierName,
                                    tier2 = pc.SecondaryTierName
                                }
                            ).ToList();

            groups_T.ForEach(
                group => {
                    for (int i = 1; i < group.Count(); i++) {
                        AddProducts(group.ElementAt(0),group.ElementAt(i));
                    }
                } 
            );

            var  objectToReturn_R = groups_T.Select(gt => gt.First()).ToList();
            return objectToReturn_R;
        }

        private static void AddProducts(ProductCategory productCategory1, ProductCategory productCategory2)
        {
            for (int i_it = 0; i_it < productCategory1.Properties.Count; i_it++) {

                if (productCategory1.Properties[i_it].Type == CellType.String &&
                        productCategory2.Properties[i_it].Type == CellType.Numeric

                )
                {
                  //  double value2_T = productCategory2.Properties[i_it].GetValue<double>();
                    productCategory1.Properties[i_it].SetValue(0);
                }
                if (productCategory1.Properties[i_it].Type == CellType.Numeric &&
                       productCategory2.Properties[i_it].Type == CellType.String

               )
                {
                    //  double value2_T = productCategory2.Properties[i_it].GetValue<double>();
                 //   productCategory1.Properties[i_it].SetValue(0);
                }



                if (productCategory1.Properties[i_it].Type == CellType.Numeric &&
                        productCategory2.Properties[i_it].Type == CellType.Numeric   
                ) {
                    double value2_T = productCategory2.Properties[i_it].GetValue<double>(); 
                    productCategory1.Properties[i_it].AddValue(value2_T);
                }

                
            }
        }

        public static ProductCategory ToProductCategory(this ExcelRowDto rowDto)
        {
            var objectToReturn_R = new ProductCategory();

            objectToReturn_R.PrimaryTierName = rowDto.ListOfCells[0].GetValueAsString();
            objectToReturn_R.SecondaryTierName = rowDto.ListOfCells[1].GetValueAsString();
            objectToReturn_R.Properties.AddRange(rowDto.ListOfCells.Skip(2));
           
            return objectToReturn_R;
        }
    }
}
