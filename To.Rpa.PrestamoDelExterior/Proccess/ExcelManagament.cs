using MAFER.FinanciamientoDelExterior.ListConfig;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using To.AtNinjas.Util;

namespace MAFER.FinanciamientoDelExterior.Proccess
{
    public class ExcelManagament
    {
        ConfigModelProduct config { get; set; }

        public ConfigModelProduct GetDataConfig(string DirConfig)
        {
            Console.WriteLine("+++++++++++++++:");
            config = new ConfigModelProduct();
            config.ListConfigCatalog = new List<ByProduct>();
            var file = new FileInfo(DirConfig);

            try
            {
                using (var package = new ExcelPackage(file))
                {
                    var workbook = package.Workbook;

                    //SubProduct
                    var worksheetSubProducto = workbook.Worksheets[4];
                    var totalRowSubProducto = worksheetSubProducto.Dimension.End.Row;

                    for (int i = 2; i <= totalRowSubProducto; i++)
                    {
                        ByProduct pd = new ByProduct();
                        pd.SizeCustomerType = GetProductValueCell(worksheetSubProducto, i, 1);
                        pd.Product_Type_Fexp = GetProductValueCell(worksheetSubProducto, i, 2);
                        pd.Product_Type_Fimp = GetProductValueCell(worksheetSubProducto, i, 3);

                        if (!pd.SizeCustomerType.ToString().Trim().Equals(""))
                        {
                            config.ListConfigCatalog.Add(pd);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Methods.LogProceso(e.ToString());
            }
            return config;
        }

        public string GetProductValueCell(ExcelWorksheet wks, int row, int col)
        {
            var cell = wks.Cells[row, col];
            if (cell.Merge)
            {
                var mergedId = wks.MergedCells[row, col];
                return wks.Cells[mergedId].First().Value.ToString();
            }
            else
            {
                if (cell.Value != null)
                    return cell.Value.ToString();
                else
                    return "";
            }



        }


    }

}