using System;
using Aspose.Cells;

namespace AsposeCellsXpsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells XPS conversion demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Initialize XPS save options
            XpsSaveOptions saveOptions = new XpsSaveOptions();

            // Configure options for high‑quality printing
            saveOptions.OnePagePerSheet = true;          // Render each sheet on a single page
            saveOptions.DefaultFont = "Arial";           // Default font for Unicode characters
            saveOptions.CheckFontCompatibility = true;  // Ensure font compatibility
            saveOptions.CheckWorkbookDefaultFont = true;
            saveOptions.PageIndex = 0;                   // Start from first page
            saveOptions.PageCount = 1;                   // Number of pages to save (adjust as needed)

            // Save the workbook as an XPS document using the specified options
            workbook.Save("WorkbookOutput.xps", saveOptions);

            Console.WriteLine("Workbook successfully saved as XPS.");
        }
    }
}