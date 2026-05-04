using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Loading;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DBF file (which contains JSON‑like data)
            string dbfPath = "source.dbf";

            // Ensure the DBF file exists; create a sample if it does not
            if (!File.Exists(dbfPath))
            {
                var sampleWb = new Workbook();
                var ws = sampleWb.Worksheets[0];
                ws.Cells["A1"].PutValue("ID");
                ws.Cells["B1"].PutValue("Name");
                ws.Cells["A2"].PutValue(1);
                ws.Cells["B2"].PutValue("John");
                sampleWb.Save(dbfPath, SaveFormat.Dbf);
            }

            // Desired output Excel file path
            string excelPath = "output.xlsx";

            // Load the DBF file using DbfLoadOptions (preserves field mappings)
            var loadOptions = new DbfLoadOptions();
            var workbook = new Workbook(dbfPath, loadOptions);

            // Save the loaded workbook as XLSX (Excel) format
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed: '{dbfPath}' → '{excelPath}'");
        }
    }
}