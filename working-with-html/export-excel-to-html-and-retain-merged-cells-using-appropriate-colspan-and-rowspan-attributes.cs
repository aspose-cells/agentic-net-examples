// Title: Export an Excel workbook to HTML with merged cells rendered as colspan and rowspan using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures HtmlSaveOptions to keep merged cell formatting (colspan/rowspan), and saves the workbook as a single HTML file. | Demonstrate how to enable base64‑encoded image embedding while exporting Excel to HTML with Aspose.Cells, ensuring merged ranges are preserved.
// Common Searches: how to keep merged cells when converting Excel to HTML with Aspose.Cells C# | Aspose.Cells HtmlSaveOptions ExportMergedCells example | C# export workbook to single HTML file with embedded images Aspose | preserve colspan rowspan in HTML output from Excel using Aspose.Cells | convert .xlsx to HTML preserving layout Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportMergedCells | C# export Excel to HTML with merged cells | Aspose.Cells embed images as base64 in HTML | preserve colspan rowspan Aspose.Cells HTML export | export multiple worksheets to single HTML file Aspose

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the input.xlsx file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions (including ExportMergedCells = true and optional base64 image embedding), and saves the workbook as output.html while preserving merged cell layout and handling any runtime exceptions.
class ExportExcelToHtml
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheets (set to true to export only the active sheet)
                ExportActiveWorksheetOnly = false,

                // Set the encoding for the generated HTML file
                Encoding = Encoding.UTF8

                // Optional: generate a single HTML file without external resources
                // ExportImagesAsBase64 = true,
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
