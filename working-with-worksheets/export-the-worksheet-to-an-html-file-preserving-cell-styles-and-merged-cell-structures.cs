// Title: Export a single worksheet to HTML with styles, merged cells, gridlines, and embedded images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, selects a specific worksheet, and saves it as an HTML file while preserving cell formatting, merged ranges, gridlines, and embedding images as Base64 with Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to export only the active sheet and keep default cell styles and merged‑cell structures in the generated HTML.
// Common Searches: Aspose.Cells .NET export only active worksheet to HTML with formatting and merged cells | How to keep gridlines and embed images when saving an Excel sheet as HTML using Aspose.Cells | C# example for saving an Excel worksheet as HTML preserving cell styles and merged ranges
// Tags: export worksheet to HTML Aspose.Cells | retain formatting in HTML export Aspose.Cells | merged cells support Aspose.Cells HTML | base64 image embedding Aspose.Cells HTML | gridlines in HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads (or creates) an Excel workbook, activates the first worksheet, configures HtmlSaveOptions to export only that sheet with gridlines, embedded Base64 images, and default style/merged‑cell preservation, and saves the result as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample");
                tempWb.Save(inputPath);
            }

            // Load the workbook from the file.
            var workbook = new Workbook(inputPath);

            // Choose the worksheet to export (first worksheet in this example).
            var worksheet = workbook.Worksheets[0];
            workbook.Worksheets.ActiveSheetIndex = worksheet.Index; // make it the active sheet

            // Configure HTML save options.
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true, // export only the selected sheet
                ExportGridLines = true,            // keep grid lines
                ExportImagesAsBase64 = true        // embed images directly
                // ExportCellStyles and ExportMergedCells are enabled by default in recent versions.
            };

            // Save the selected worksheet as an HTML file.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
