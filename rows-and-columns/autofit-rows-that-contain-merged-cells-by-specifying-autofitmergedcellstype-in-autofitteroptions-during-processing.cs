// Title: C# Example: Auto‑fit Rows with Merged Cells Using AutoFitterOptions in Aspose.Cells for .NET
// Description: Demonstrates how to merge cells, enable text wrapping, configure AutoFitterOptions (AutoFitMergedCellsType.EachLine and AutoFitWrappedTextType.Paragraph), and call worksheet.AutoFitRows(options) so the row height automatically adapts to wrapped text inside merged ranges before saving the workbook.
// Keywords: Aspose.Cells AutoFitRows C# | AutoFitterOptions merged cells | AutoFitMergedCellsType EachLine | AutoFitWrappedTextType Paragraph | .NET Excel row height auto fit | wrap text merged cells Aspose | Excel row auto‑adjust C# | GitHub Aspose.Cells sample
// Common Searches: auto fit rows with merged cells Aspose.Cells .NET | AutoFitMergedCellsType each line example | how to include wrapped text when auto‑fitting rows Aspose | C# code for AutoFitRows using AutoFitterOptions | Aspose.Cells merge cells and adjust row height
// Developer Intent: Automatically adjust row heights to display all wrapped text inside merged cells by using AutoFitterOptions with appropriate settings.
// Use Cases: Create a multi‑column title that spans rows and automatically expands to fit long wrapped headings. | Design invoice or report templates where header cells are merged and description fields vary in length. | Export data sets containing multi‑line comments in merged cells, ensuring rows resize without manual intervention.
// AI Prompts: Show a C# snippet that merges A1:B3, enables text wrap, sets AutoFitMergedCellsType.EachLine and AutoFitWrappedTextType.Paragraph, then auto‑fits rows with Aspose.Cells. | Explain how AutoFitMergedCellsType.EachLine differs from AutoFitMergedCellsType.AllLines when auto‑fitting rows in merged cells. | Provide step‑by‑step instructions for using AutoFitterOptions to auto‑fit rows containing wrapped text in merged cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to merge cells, enable text wrapping, configure AutoFitterOptions (AutoFitMergedCellsType.EachLine and AutoFitWrappedTextType.Paragraph), and call worksheet.AutoFitRows(options) so the row height automatically adapts to wrapped text inside merged ranges before saving the workbook.
    public class AutoFitRowsWithMergedCellsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a long text into a cell
                worksheet.Cells["A1"].PutValue(
                    "This is a sample text for merged cells auto‑fit demonstration. It should wrap and cause the row height to adjust based on the merged cell settings.");

                // Merge a range of cells (A1:B3)
                worksheet.Cells.Merge(0, 0, 3, 2);

                // Enable text wrapping for the merged cell
                Style style = worksheet.Cells["A1"].GetStyle();
                style.IsTextWrapped = true;
                worksheet.Cells["A1"].SetStyle(style);

                // Configure AutoFitterOptions to consider merged cells (each line) and wrapped text (paragraph)
                AutoFitterOptions options = new AutoFitterOptions
                {
                    AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                    AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
                };

                // Auto‑fit rows using the specified options (feature rule: AutoFitRows(AutoFitterOptions))
                worksheet.AutoFitRows(options);

                // Save the workbook (lifecycle rule: save)
                string outputPath = "AutoFitRowsWithMergedCellsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            AutoFitRowsWithMergedCellsDemo.Run();
        }
    }
}
