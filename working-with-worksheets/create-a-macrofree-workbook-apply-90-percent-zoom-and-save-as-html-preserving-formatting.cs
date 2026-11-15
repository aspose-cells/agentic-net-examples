// Title: Create a macro‑free workbook, set worksheet zoom to 90 % and save as HTML with formatting preserved using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet zoom level to 90 percent | How to export a workbook to HTML with embedded images using Aspose.Cells | Saving a macro‑free Excel file as HTML while keeping cell styles in .NET | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly false example | Create workbook without macros and convert to HTML in C#
// Tags: macro‑free workbook creation Aspose.Cells | worksheet zoom property Aspose.Cells | HTML export with formatting Aspose.Cells | ExportImagesAsBase64 option Aspose.Cells | export entire workbook to HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new macro‑free Workbook, sets the first worksheet's Zoom to 90 %, configures HtmlSaveOptions to export the whole workbook and embed images as Base64, ensures the output directory exists, and saves the result as an HTML file while handling possible exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new macro‑free workbook
            Workbook workbook = new Workbook();

            // Set the zoom level of the first worksheet to 90%
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Zoom = 90;

            // Configure HTML save options to preserve formatting
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export the entire workbook (not only the active sheet)
                ExportActiveWorksheetOnly = false,
                // Embed images (including charts) as Base64 to keep formatting
                ExportImagesAsBase64 = true
                // Note: IsFullHtml and ExportCellStyles are not available in the current API version
            };

            // Define output file path
            string outputPath = "output.html";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
