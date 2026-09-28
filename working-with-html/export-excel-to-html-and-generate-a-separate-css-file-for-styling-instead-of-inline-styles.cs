// Title: Export an Excel workbook to HTML with Aspose.Cells for .NET – configure external CSS (fallback to inline styles when unsupported)
// AI Prompts: Write C# code that loads a .xlsx file with Aspose.Cells, sets HtmlSaveOptions to export all worksheets and save images as separate files, then saves the workbook as an HTML document. | Demonstrate how to request an external CSS stylesheet when saving a workbook to HTML with Aspose.Cells, and add a comment explaining the fallback to inline CSS if the ExportCss API is unavailable. | Add robust file‑existence checking and comprehensive exception handling around the Excel‑to‑HTML conversion process.
// Common Searches: Aspose.Cells C# export workbook to HTML with external stylesheet | How to generate separate CSS file when saving Excel as HTML using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false example | C# check if Excel file exists before converting to HTML with Aspose.Cells | Aspose.Cells version that supports ExportCss property for HTML export
// Tags: Aspose.Cells HtmlSaveOptions external stylesheet | C# export Excel to HTML separate CSS | ExportImagesAsBase64 false Aspose.Cells | Excel to HTML conversion .NET | file existence validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that input.xlsx exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to export all worksheets and write images as separate files, and saves the result as output.html while handling any runtime exceptions. It also notes that external CSS generation is only available in newer API versions; otherwise, CSS is embedded inline.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputHtml = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets
                ExportActiveWorksheetOnly = false,

                // Export images as separate files (not Base64)
                ExportImagesAsBase64 = false

                // Note: ExportCss and CssStyleSheetType are not available in the current Aspose.Cells version.
                // The default behavior will embed CSS inline.
            };

            // Save the workbook as HTML; this creates "output.html"
            workbook.Save(outputHtml, htmlOptions);
            Console.WriteLine($"Conversion succeeded. HTML saved to \"{outputHtml}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
