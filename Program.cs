using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace EpPlusDemo.NetCore
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting...");

            List<string> values = new List<string>();

            byte[] file = File.ReadAllBytes(@".\Source\Template.xlsx");
            using (MemoryStream ms = new MemoryStream(file))
            {
                using (ExcelPackage package = new ExcelPackage(ms))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                        Console.WriteLine("Workbook contains no sheets");
                    else
                    {
                        foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                        {
                            int row = worksheet.Names["start"].Start.Row;
                            int col = worksheet.Names["start"].Start.Column;
                            object val = worksheet.Cells[row, col].Value;
                            while(val != null)
                            {
                                values.Add(val.ToString());
                                row++;
                                val = worksheet.Cells[row, col].Value;
                            }
                        }
                    }
                }
            }
        }
    }
}