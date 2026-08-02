// Title: Fit all worksheet columns on one printed page with Aspose.Cells for .NET
// Description: Demonstrates how to configure a workbook’s PageSetup in Aspose.Cells so the printed output fits every column on a single page (FitToPagesWide = 1, FitToPagesTall = 0) while disabling percentage scaling to avoid distortion.
// Keywords: Aspose.Cells print scaling | FitToPagesWide | FitToPagesTall | IsPercentScale false | C# Excel page setup | fit columns one page | worksheet print layout .NET | Aspose.Cells pagination
// Common Searches: Aspose.Cells fit columns to one page C# | set worksheet page setup FitToPagesWide Aspose | disable percent scaling Aspose.Cells | print Excel sheet without horizontal scroll .NET | how to fit all columns on a single printed page using Aspose
// Developer Intent: Apply page‑setup properties to print every column on one page without using percentage scaling.
// Use Cases: Creating printable reports that must stay within a single page width. | Generating invoices or receipts where column layout should not wrap across pages. | Exporting data tables for PDF conversion while preserving column alignment on one page.
// AI Prompts: Provide C# code that sets FitToPagesWide = 1, FitToPagesTall = 0, and IsPercentScale = false in Aspose.Cells. | Explain the effect of disabling percentage scaling when using FitToPages settings in Aspose.Cells. | Show how to configure Aspose.Cells page setup to fit all columns on one page while allowing rows to span multiple pages.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to configure a workbook’s PageSetup in Aspose.Cells so the printed output fits every column on a single page (FitToPagesWide = 1, FitToPagesTall = 0) while disabling percentage scaling to avoid distortion.
    public class FitAllColumnsOnePage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // (Optional) Add some sample data to demonstrate the effect
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Set the page setup to fit all columns on a single page.
                // FitToPagesWide = 1 means one page wide.
                // FitToPagesTall = 0 lets the height adjust automatically.
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0;

                // Ensure scaling is based on FitToPages settings, not percent scaling.
                worksheet.PageSetup.IsPercentScale = false;

                // Save the workbook (adjust the path/format as needed)
                workbook.Save("FitAllColumnsOnePage.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            FitAllColumnsOnePage.Run();
        }
    }
}
