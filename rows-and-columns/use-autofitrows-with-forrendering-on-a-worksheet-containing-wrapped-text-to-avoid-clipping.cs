// Title: C# AutoFitRows + ForRendering – keep wrapped text visible in Aspose.Cells
// Description: The sample creates a workbook, writes a long string to A1, turns on text wrap, configures AutoFitterOptions with ForRendering enabled, runs Worksheet.AutoFitRows, prints the computed row height, and saves the file, ensuring the wrapped paragraph is not cut off in the rendered output.
// Keywords: Aspose.Cells | AutoFitRows | ForRendering | C# | text wrap | row height adjustment | Excel export | PDF rendering | prevent text clipping | AutoFitterOptions
// Common Searches: Aspose.Cells AutoFitRows rendering mode C# | prevent clipped wrapped cells Aspose.Cells | AutoFitterOptions ForRendering usage example | auto adjust row height for wrapped text | C# code to fit rows for PDF export
// Developer Intent: Determine the proper row height for wrapped cells when generating Excel or PDF output.
// Use Cases: Generate reports where cells contain paragraphs and must appear fully in on‑screen or printed PDFs. | Programmatically size rows before exporting large data sets to maintain layout integrity. | Create templates that automatically adapt row heights after filling dynamic content.
// AI Prompts: Write a C# snippet that sets text wrapping, enables AutoFitterOptions.ForRendering, auto‑fits rows, and returns the new heights. | Explain the impact of the ForRendering flag on AutoFitRows and when to use it for PDF generation. | Show how to combine AutoFitRows with column auto‑fit to produce a perfectly sized worksheet in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AutoFitRowsForRenderingDemo
{
    // The sample creates a workbook, writes a long string to A1, turns on text wrap, configures AutoFitterOptions with ForRendering enabled, runs Worksheet.AutoFitRows, prints the computed row height, and saves the file, ensuring the wrapped paragraph is not cut off in the rendered output.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into cell A1
            worksheet.Cells["A1"].PutValue("This is a very long text that should be wrapped and auto‑fitted for rendering purposes. It contains enough characters to require multiple lines when wrapped.");

            // Enable text wrapping for the cell
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Create AutoFitterOptions and set ForRendering to true
            AutoFitterOptions options = new AutoFitterOptions
            {
                ForRendering = true
            };

            // Auto‑fit the row containing the wrapped text with rendering considerations
            worksheet.AutoFitRows(options);

            // Output the resulting row height in points (or pixels if needed)
            Console.WriteLine("Row 0 height after AutoFitRows (points): " + worksheet.Cells.GetRowHeight(0));

            // Save the workbook to a file
            workbook.Save("AutoFitRowsForRenderingDemo.xlsx");
        }
    }
}
