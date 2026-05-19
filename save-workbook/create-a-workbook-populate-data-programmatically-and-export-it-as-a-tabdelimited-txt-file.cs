using System;
using Aspose.Cells;

namespace AsposeCellsTxtExportExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate data programmatically
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(999.99);
            worksheet.Cells["A3"].PutValue("Phone");
            worksheet.Cells["B3"].PutValue(699.99);
            worksheet.Cells["A4"].PutValue("Tablet");
            worksheet.Cells["B4"].PutValue(399.99);

            // Configure text save options for tab‑delimited output (rule usage)
            TxtSaveOptions txtOptions = new TxtSaveOptions();
            txtOptions.Separator = '\t';               // Tab character as delimiter
            txtOptions.Encoding = System.Text.Encoding.UTF8;
            txtOptions.ExportAllSheets = false;        // Export only the active sheet

            // Save the workbook as a tab‑delimited TXT file (lifecycle rule)
            string outputPath = "ExportedData.txt";
            workbook.Save(outputPath, txtOptions);

            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
    }
}