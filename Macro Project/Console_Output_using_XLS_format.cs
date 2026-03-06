using System;
using Aspose.Cells;

namespace AsposeCellsConsoleXlsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue(999.99);
            worksheet.Cells["A3"].PutValue("Phone");
            worksheet.Cells["B3"].PutValue(599.99);

            // Create XlsSaveOptions (using the provided constructor rule)
            XlsSaveOptions saveOptions = new XlsSaveOptions();

            // Set a few options (optional, demonstrates usage of properties)
            saveOptions.MatchColor = true;      // Match font colors to the 56‑color palette
            saveOptions.IsTemplate = false;    // Not saving as a template (obsolete property)

            // Save the workbook as an Excel 97‑2003 .xls file (lifecycle: save)
            string outputPath = "SampleOutput.xls";
            workbook.Save(outputPath, saveOptions);

            // Console output to confirm success
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}