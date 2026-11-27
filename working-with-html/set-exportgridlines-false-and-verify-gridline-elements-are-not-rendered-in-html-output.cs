// Title: Export an Excel worksheet to HTML without grid lines using Aspose.Cells for .NET and verify the result
// Common Searches: Aspose.Cells C# export worksheet to HTML without grid lines | how to set ExportGridLines false in HtmlSaveOptions and check output | C# read generated HTML file to ensure no gridline CSS classes | validate that Aspose.Cells HTML export does not include border styles | save Excel as HTML without gridlines and confirm via code
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines false | C# export Excel to HTML without grid lines | verify HTML output for gridline CSS Aspose.Cells | disable grid lines in HTML export Aspose.Cells | read and validate generated HTML file C#

using System;
using System.IO;
using Aspose.Cells;

namespace ExportGridLinesDemo
{
    // The example creates a workbook, populates it with sample data, and saves it as an HTML file using HtmlSaveOptions with ExportGridLines set to false. It then reads the generated HTML and checks for any gridline-related CSS classes or border styles, reporting whether the verification passed or failed.
    class Program
    {
        static void Main(string[] args)
        {
            // Path for the temporary HTML file
            string htmlPath = Path.Combine(Path.GetTempPath(), "ExportGridLinesDemo.html");

            // 1. Create a workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill a small table
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(1.20);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(0.80);
            cells["A4"].PutValue("Cherry");
            cells["B4"].PutValue(2.50);

            // 2. Configure HTML save options: ExportGridLines = false
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = false,          // Ensure grid lines are NOT exported
                ExportHeadings = true,            // Export row/column headings for readability
                ExportActiveWorksheetOnly = true  // Only the active sheet is saved
            };

            // 3. Save the workbook to HTML using the configured options
            try
            {
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"HTML file saved to: {htmlPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during HTML save: {ex.Message}");
                return;
            }

            // 4. Verify that the generated HTML does NOT contain grid line definitions
            //    Grid lines are typically rendered via CSS classes like "gridlines" or inline styles with "border"
            bool gridLinesFound = false;
            try
            {
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine("HTML file was not found for verification.");
                    return;
                }

                string htmlContent = File.ReadAllText(htmlPath);

                // Simple heuristics:
                // - Look for a CSS class named "gridlines"
                // - Look for any occurrence of "border:" in style attributes
                if (htmlContent.IndexOf("gridlines", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    htmlContent.IndexOf("border:", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    gridLinesFound = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading HTML file: {ex.Message}");
                return;
            }

            // 5. Output verification result
            if (gridLinesFound)
            {
                Console.WriteLine("Verification FAILED: Grid line elements were found in the HTML output.");
            }
            else
            {
                Console.WriteLine("Verification PASSED: No grid line elements are present in the HTML output.");
            }

            // Optional: Clean up the temporary file (uncomment if you want automatic deletion)
            // try { File.Delete(htmlPath); } catch { /* ignore */ }
        }
    }
}
