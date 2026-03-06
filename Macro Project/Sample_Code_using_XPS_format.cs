using System;
using Aspose.Cells;

namespace AsposeCellsXpsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells XPS Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);

            // Create XpsSaveOptions using the default constructor
            XpsSaveOptions saveOptions = new XpsSaveOptions();

            // Configure the XPS save options
            saveOptions.OnePagePerSheet = true;               // Render each sheet on a single page
            saveOptions.DefaultFont = "Arial";                // Default font for Unicode characters
            saveOptions.CheckWorkbookDefaultFont = true;      // Use workbook default font when needed
            saveOptions.CheckFontCompatibility = true;        // Verify font compatibility
            saveOptions.PageIndex = 0;                        // Start from the first page
            saveOptions.PageCount = 1;                        // Save only one page (adjust as needed)

            // Save the workbook as an XPS file using the configured options
            workbook.Save("DemoOutput.xps", saveOptions);
        }
    }
}