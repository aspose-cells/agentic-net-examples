// Title: AutoFitRows with ForRendering to prevent wrapped‑text clipping in Aspose.Cells for .NET
// Description: Demonstrates how to enable text wrapping, set AutoFitterOptions.ForRendering, and call Worksheet.AutoFitRows so that long wrapped strings are fully visible and not clipped. The sample prints the row height in pixels and saves the workbook, ideal for rendering to Excel, PDF, or images.
// Keywords: Aspose.Cells | .NET | C# | AutoFitRows | ForRendering | AutoFitterOptions | wrapped text | row height | pixel height | Excel export | PDF rendering | image export
// Common Searches: Aspose.Cells AutoFitRows ForRendering example | auto‑fit rows with wrapped text .NET | prevent text clipping when rendering Aspose.Cells | get row height in pixels after AutoFitRows | auto‑fit rows for PDF export Aspose.Cells
// Developer Intent: Automatically adjust row height for wrapped content during rendering to ensure no clipping.
// Use Cases: Fit a single row that contains a long wrapped string before saving to Excel or PDF. | Retrieve the exact pixel height of a row after AutoFitRows for UI layout calculations. | Apply ForRendering mode when exporting workbooks to image formats so wrapped text displays completely.
// AI Prompts: Write C# code that iterates over a range of rows, applies AutoFitRows with AutoFitterOptions.ForRendering, and logs each row's pixel height. | Explain the difference between the default AutoFitRows behavior and the ForRendering option, and recommend scenarios for each. | Generate a complete Aspose.Cells example that auto‑fits rows with wrapped text and then exports the workbook to PDF without clipping.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable text wrapping, set AutoFitterOptions.ForRendering, and call Worksheet.AutoFitRows so that long wrapped strings are fully visible and not clipped. The sample prints the row height in pixels and saves the workbook, ideal for rendering to Excel, PDF, or images.
    public class AutoFitRowsForRenderingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a long text into cell A1
                worksheet.Cells["A1"].PutValue(
                    "This is a very long text that should be wrapped and auto‑fitted for rendering purposes. It contains enough characters to require multiple lines when wrapped.");

                // Enable text wrapping for the cell
                Style style = worksheet.Cells["A1"].GetStyle();
                style.IsTextWrapped = true;
                worksheet.Cells["A1"].SetStyle(style);

                // Create AutoFitterOptions and enable rendering mode
                AutoFitterOptions options = new AutoFitterOptions
                {
                    ForRendering = true
                };

                // Auto‑fit the rows using the options (row 0 only in this case)
                worksheet.AutoFitRows(0, 0, options);

                // Output the resulting row height in pixels
                Console.WriteLine("Row height after AutoFitRows (ForRendering = true): " +
                                  worksheet.Cells.GetRowHeightPixel(0));

                // Save the workbook to verify the result
                string outputPath = "AutoFitRowsForRenderingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
