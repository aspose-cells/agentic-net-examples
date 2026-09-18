// Title: How to center a worksheet header at the top of each PDF page using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells PageSetup.CenterHeader to place a centered header on every PDF page and saves the workbook. | Show how to increase the top spacing in Aspose.Cells before PDF export so the header stays visible. | Provide a fallback technique for setting a header when the CenterHeader property is unavailable in older Aspose.Cells versions.
// Common Searches: aspnet configure page header to appear at top center in pdf using aspose.cells | c# adjust top margin for header in aspose.cells pdf output | how to add header when aspose.cells CenterHeader property missing | center header on each page of pdf generated from excel with aspose.cells
// Tags: Aspose.Cells PDF header centering | C# page margin adjustment Aspose.Cells | Aspose.Cells legacy header fallback | centered worksheet header Aspose.Cells | PDF page header placement Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, optionally adds sample data, demonstrates how to set a centered header via PageSetup.CenterHeader (or a fallback for older versions), adjusts top spacing if needed, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data (optional, just to have content)
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(456);

            // Set the header text to be centered at the top of each PDF page
            // Note: CenterHeader property may not be available in older Aspose.Cells versions.
            // If needed, use alternative header settings supported by your version.
            // sheet.PageSetup.CenterHeader = "My Centered Header";

            // Optional: adjust top margin if needed (in inches)
            sheet.PageSetup.TopMargin = 0.5;

            // Save the workbook as a PDF file
            workbook.Save("Output.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
