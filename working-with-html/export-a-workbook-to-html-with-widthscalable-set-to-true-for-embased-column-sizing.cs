// Title: Export Workbook to HTML with WidthScalable (em‑based column sizing) using Aspose.Cells C#
// Description: Demonstrates how to create a workbook, add data, enable the WidthScalable option in HtmlSaveOptions to generate column widths in em units, and save the result as a responsive HTML file.
// Keywords: Aspose.Cells | HtmlSaveOptions | WidthScalable | C# export to HTML | em based column width | responsive Excel HTML | scalable column sizing | Excel to HTML conversion
// Common Searches: Aspose.Cells WidthScalable true example | export Excel to HTML with em units | C# save workbook as responsive HTML | how to use HtmlSaveOptions WidthScalable | HTML column scaling Aspose.Cells
// Developer Intent: Create an HTML representation of an Excel workbook where column widths are expressed in scalable em units instead of fixed pixels.
// Use Cases: Display Excel data on web pages that adapt to different screen sizes. | Generate email‑friendly HTML reports with proportionate column widths. | Build responsive dashboards that preserve Excel layout without pixel‑based constraints.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as HTML using WidthScalable=true and add custom CSS for table styling. | Explain the effect of the WidthScalable property on column width calculation and how to influence the resulting em values. | Provide a step‑by‑step guide to export multiple worksheets into a single responsive HTML file while keeping column widths scalable.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add data, enable the WidthScalable option in HtmlSaveOptions to generate column widths in em units, and save the result as a responsive HTML file.
    public class ExportWorkbookToHtmlWithWidthScalable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Column A");
                sheet.Cells["B1"].PutValue("Column B");
                sheet.Cells["A2"].PutValue("Short");
                sheet.Cells["B2"].PutValue("A much longer piece of text that would normally be truncated");

                // Configure HTML save options to use scalable column widths (em units)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    WidthScalable = true // Enable em‑based column sizing
                };

                // Save the workbook as an HTML file using the configured options
                string outputPath = "Workbook_With_WidthScalable.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
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
            ExportWorkbookToHtmlWithWidthScalable.Run();
        }
    }
}
