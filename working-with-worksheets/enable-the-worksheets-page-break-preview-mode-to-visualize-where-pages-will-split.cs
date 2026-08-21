// Title: Enable Page Break Preview and Set Zoom for a Worksheet with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, activate the Page Break Preview mode, adjust the worksheet zoom level, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells page break preview | IsPageBreakPreview .NET | worksheet zoom Aspose.Cells | enable page break preview C# | Aspose.Cells save workbook | Excel page layout view programmatically
// Common Searches: how to turn on page break preview in Aspose.Cells | set worksheet zoom while enabling page break preview .NET | Aspose.Cells enable page break preview and save workbook | C# code for page break preview mode in Excel file | Aspose.Cells display page breaks before printing
// Developer Intent: Activate Page Break Preview on a worksheet, set a specific zoom factor, and persist the workbook.
// Use Cases: Prepare a spreadsheet that shows exact page breaks for printing or PDF export. | Create a report with a fixed zoom level so all users see the same layout when opening the file. | Generate a printable preview where pagination is visible to ensure content fits on intended pages.
// AI Prompts: Show code to enable page break preview for every worksheet in an existing workbook using Aspose.Cells for .NET. | Provide a C# example that toggles IsPageBreakPreview based on a condition and adjusts the zoom accordingly. | Explain the relationship between IsPageBreakPreview, Excel's Page Layout view, and printed page breaks.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, activate the Page Break Preview mode, adjust the worksheet zoom level, and save the file using Aspose.Cells for .NET.
    public class EnablePageBreakPreview
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable Page Break Preview mode
                worksheet.IsPageBreakPreview = true;

                // Optional: set zoom to 100% for clearer view
                worksheet.Zoom = 100;

                // Output the current settings
                Console.WriteLine("IsPageBreakPreview: " + worksheet.IsPageBreakPreview);
                Console.WriteLine("Zoom: " + worksheet.Zoom);

                // Save the workbook
                string outputPath = "EnablePageBreakPreview_output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnablePageBreakPreview.Run();
        }
    }
}
