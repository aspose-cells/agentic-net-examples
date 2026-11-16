// Title: Export Excel workbook to HTML with functional hyperlinks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export Excel to HTML with clickable hyperlinks | How to keep Excel hyperlinks functional when saving as HTML using Aspose.Cells | Save workbook as self‑contained HTML file with embedded images Aspose.Cells C# | Aspose.Cells HtmlSaveOptions settings for preserving links and images | Convert .xlsx to HTML preserving formatting and links in .NET
// Tags: Aspose.Cells HtmlSaveOptions export HTML | maintain hyperlink functionality in HTML | embed images as Base64 in HTML | UTF-8 encoding for HTML export | export all worksheets to single HTML file

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an existing .xlsx file with Aspose.Cells, configures HtmlSaveOptions to export all worksheets, use UTF-8 encoding, and embed images as Base64 strings, then saves the workbook as an HTML file. Hyperlinks from the original workbook remain clickable in the browser, and the code includes basic file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
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
                // Export all worksheets (set to true to export only the active sheet)
                ExportActiveWorksheetOnly = false,

                // Use UTF-8 encoding for proper character rendering
                Encoding = Encoding.UTF8,

                // Embed images directly into the HTML as Base64 strings (optional but keeps files self‑contained)
                ExportImagesAsBase64 = true
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
