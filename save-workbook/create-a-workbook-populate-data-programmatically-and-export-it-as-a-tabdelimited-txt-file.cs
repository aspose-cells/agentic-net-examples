using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTxtExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data programmatically
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue(999.99);
            sheet.Cells["A3"].PutValue("Phone");
            sheet.Cells["B3"].PutValue(599.99);
            sheet.Cells["A4"].PutValue("Tablet");
            sheet.Cells["B4"].PutValue(399.99);

            // Configure text save options for tab‑delimited output
            TxtSaveOptions txtOptions = new TxtSaveOptions
            {
                // Use tab character as the separator
                Separator = '\t',
                // Ensure UTF‑8 encoding for proper character handling
                Encoding = Encoding.UTF8
            };

            // Save the workbook as a tab‑delimited TXT file (lifecycle rule: save)
            workbook.Save("ExportedData.txt", txtOptions);

            Console.WriteLine("Workbook has been exported as a tab‑delimited TXT file.");
        }
    }
}