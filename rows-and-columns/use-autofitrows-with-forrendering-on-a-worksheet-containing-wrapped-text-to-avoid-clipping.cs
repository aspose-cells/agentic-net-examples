// Title: Auto-fit a worksheet row with wrapped text for rendering using Aspose.Cells AutoFitterOptions in C#
// AI Prompts: Generate C# code that creates a workbook, wraps long text in a cell, and calls worksheet.AutoFitRows with AutoFitterOptions.ForRendering to prevent clipping. | Write C# to retrieve the pixel height of a specific row after applying AutoFitRows with the rendering option. | Provide C# that saves the workbook after auto‑fitting rows for rendering and prints the resulting row height.
// Common Searches: Aspose.Cells C# AutoFitRows ForRendering option to avoid text clipping | How to get row height in pixels after AutoFitRows in Aspose.Cells | Wrap text and auto‑fit rows for PDF export using Aspose.Cells C# example | AutoFitRows rendering option sample code C#
// Tags: AutoFitRows with AutoFitterOptions rendering | wrapped text row height pixels Aspose.Cells | C# Aspose.Cells row auto‑fit for PDF export | prevent text clipping Excel Aspose.Cells | AutoFitterOptions ForRendering usage C#

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, inserts a long wrapped text into cell A1, enables text wrapping, configures AutoFitterOptions with ForRendering=true, auto‑fits the first row, outputs the row height in pixels, and saves the file as AutoFitRowsForRendering.xlsx.
    public class AutoFitRowsForRenderingDemo
    {
        // Entry point for the console application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into a cell and enable text wrapping
            worksheet.Cells["A1"].PutValue(
                "This is a very long piece of text that should be wrapped and auto‑fitted for rendering purposes. It contains enough characters to require multiple lines when wrapped.");
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Configure AutoFitterOptions to fit for rendering
            AutoFitterOptions options = new AutoFitterOptions
            {
                ForRendering = true
            };

            // Auto‑fit the first row using the rendering option
            worksheet.AutoFitRows(0, 0, options);

            // Output the resulting row height in pixels
            Console.WriteLine("Row height after AutoFitRows (pixels): " + worksheet.Cells.GetRowHeightPixel(0));

            // Save the workbook
            workbook.Save("AutoFitRowsForRendering.xlsx");
        }
    }
}
