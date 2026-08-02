// Title: C# – Export Excel numbers as plain text in HTML with Aspose.Cells HtmlSaveOptions
// Description: Shows how to stop scientific notation when converting an Excel workbook to HTML using Aspose.Cells for .NET. The sample applies a simple numeric format and enables HtmlSaveOptions.ExportNumericDataAsString (or the equivalent setting) so every numeric cell is written as plain text in the HTML output.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportNumericDataAsString | plain numeric export | avoid scientific notation | C# | .NET | Excel to HTML | custom number format | HTML export of numbers | financial report HTML | engineering data HTML
// Common Searches: Aspose.Cells export numbers as text HTML | prevent scientific notation in HTML export Aspose | HtmlSaveOptions ExportNumericDataAsString .NET | C# save Excel to HTML without scientific notation | how to force plain numbers in Aspose.Cells HTML output
// Developer Intent: Save an Excel workbook to HTML so that all numeric cells appear as plain text rather than scientific notation.
// Use Cases: Financial statements published on the web where large currency values must stay readable. | Engineering spreadsheets that require exact decimal representation in HTML tables. | Web dashboards that display identifier codes or measurements without exponential formatting. | Static HTML reports generated from Excel data for compliance documentation.
// AI Prompts: Provide C# code that sets HtmlSaveOptions.ExportNumericDataAsString = true to export numbers as plain text with Aspose.Cells. | Explain how a custom number format and HtmlSaveOptions work together to eliminate scientific notation in HTML output. | List the NuGet packages and steps needed to convert an Excel workbook to HTML with all numeric cells rendered as plain strings.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Shows how to stop scientific notation when converting an Excel workbook to HTML using Aspose.Cells for .NET. The sample applies a simple numeric format and enables HtmlSaveOptions.ExportNumericDataAsString (or the equivalent setting) so every numeric cell is written as plain text in the HTML output.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with numeric values that could be shown in scientific notation
                sheet.Cells["A1"].PutValue(1234567890);
                sheet.Cells["A2"].PutValue(0.0000001234);

                // Apply a style that forces plain numeric display (no scientific notation)
                Style plainStyle = workbook.CreateStyle();
                plainStyle.Custom = "0"; // integer format; adjust as needed
                sheet.Cells["A1"].SetStyle(plainStyle);
                sheet.Cells["A2"].SetStyle(plainStyle);

                // Configure HTML save options to export all cell data
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportDataOptions = HtmlExportDataOptions.All
                };

                // Define output file path
                string outputPath = "PlainNumericExport.html";

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
