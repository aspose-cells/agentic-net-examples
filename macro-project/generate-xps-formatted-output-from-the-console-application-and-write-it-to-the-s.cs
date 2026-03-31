using System;
using Aspose.Cells;

namespace XpsExportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells XPS Export Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(67.89);

            // Initialize XpsSaveOptions using the default constructor
            XpsSaveOptions saveOptions = new XpsSaveOptions();

            // Optional: configure desired options
            saveOptions.OnePagePerSheet = true;          // Export each sheet as a single page
            saveOptions.DefaultFont = "Arial";           // Fallback font for Unicode characters
            saveOptions.PageIndex = 0;                   // Start from the first page
            saveOptions.PageCount = 1;                   // Export only one page (adjust as needed)

            // Define the output file path
            string outputPath = "ExportedDocument.xps";

            // Save the workbook as XPS using the specified options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved as XPS to: {outputPath}");
        }
    }
}