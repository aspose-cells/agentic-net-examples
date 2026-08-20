// Title: Aspose.Cells for .NET – Fixed Pixel Column Widths with WidthScalable = false in HTML export
// Description: Demonstrates how to set column widths in exact pixels using SetColumnWidthPixel, verify them with GetColumnWidthPixel, and generate HTML where the columns keep those fixed sizes by setting HtmlSaveOptions.WidthScalable to false. The workbook is also saved as XLSX to show that the pixel‑based widths persist in the native format.
// Keywords: Aspose.Cells column width pixel | HtmlSaveOptions WidthScalable false | C# export Excel to HTML fixed width | SetColumnWidthPixel example | disable column scaling Aspose.Cells
// Common Searches: Aspose.Cells set column width in pixels | How to disable WidthScalable in HTML export | Export Excel to HTML with fixed column sizes .NET | GetColumnWidthPixel Aspose.Cells C# | Fixed pixel column width HTML Aspose
// Developer Intent: Create HTML output where column widths are locked to specific pixel values by turning off WidthScalable.
// Use Cases: Define precise pixel widths for columns A and B, then export the sheet to HTML with non‑scalable columns. | Read back the pixel dimensions after setting them to confirm the values before saving. | Save the same workbook as XLSX to ensure the pixel‑based column settings are retained in the original file.
// AI Prompts: Generate C# code that sets column widths in pixels with Aspose.Cells and saves the workbook to HTML with WidthScalable disabled. | Explain the effect of HtmlSaveOptions.WidthScalable = false on the generated HTML and how to validate column widths. | Provide a step‑by‑step tutorial for verifying pixel column widths before and after exporting to HTML and XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to set column widths in exact pixels using SetColumnWidthPixel, verify them with GetColumnWidthPixel, and generate HTML where the columns keep those fixed sizes by setting HtmlSaveOptions.WidthScalable to false. The workbook is also saved as XLSX to show that the pixel‑based widths persist in the native format.
    public class WidthScalableFixedDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set column widths in pixels (fixed values)
                cells.SetColumnWidthPixel(0, 150); // Column A = 150 pixels
                cells.SetColumnWidthPixel(1, 200); // Column B = 200 pixels

                // Verify the pixel widths
                int col0Width = cells.GetColumnWidthPixel(0);
                int col1Width = cells.GetColumnWidthPixel(1);
                Console.WriteLine($"Column A width (pixels): {col0Width}");
                Console.WriteLine($"Column B width (pixels): {col1Width}");

                // Populate some data to visualize the widths in HTML
                sheet.Cells["A1"].PutValue("Short");
                sheet.Cells["B1"].PutValue("This is a longer text to show column width effect");

                // Configure HTML save options with WidthScalable set to false (fixed pixel widths)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    WidthScalable = false
                };

                // Save the workbook as HTML with fixed column widths
                workbook.Save("output_fixed.html", htmlOptions);

                // Also save as XLSX to confirm the widths are retained in the native format
                workbook.Save("output_fixed.xlsx");

                Console.WriteLine("Files saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WidthScalableFixedDemo.Run();
        }
    }
}
