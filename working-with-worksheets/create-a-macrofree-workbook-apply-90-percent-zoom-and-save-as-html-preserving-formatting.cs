// Title: Generate a macro‑free Excel workbook, set worksheet zoom to 90 %, and save it as HTML with formatting preserved using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create workbook without macros and export to HTML preserving layout | How to set worksheet zoom to 90 percent in Aspose.Cells before saving as HTML | Save Excel file as HTML with embedded images using Aspose.Cells .NET | Export entire workbook to HTML with original formatting using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions to embed images as Base64 and keep styling
// Tags: macro‑free workbook creation Aspose.Cells | worksheet zoom setting Aspose.Cells | HTML export preserving formatting Aspose.Cells | embed images as Base64 Aspose.Cells HTML | export entire workbook to HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new macro‑free Workbook, accesses the first worksheet, sets its PageSetup.Zoom to 90 %, configures HtmlSaveOptions to export the whole workbook with images embedded as Base64, and saves the result as an HTML file while retaining the original Excel formatting.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (macro‑free by default)
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Set the worksheet view zoom to 90%
                sheet.PageSetup.Zoom = 90;

                // Prepare HTML save options to keep the original formatting
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export the whole workbook (set to true to export only the active worksheet)
                    ExportActiveWorksheetOnly = false,
                    // Export images as Base64 strings so they are embedded in the HTML
                    ExportImagesAsBase64 = true
                    // Note: HtmlFormattingOptions property removed to avoid compatibility issues
                };

                // Save the workbook as an HTML file while preserving formatting
                string outputPath = "output.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
