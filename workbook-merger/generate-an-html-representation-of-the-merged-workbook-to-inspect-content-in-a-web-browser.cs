// Title: Generate an HTML file from a merged Excel workbook with embedded images using Aspose.Cells for .NET (C#)
// Common Searches: how to export a merged .xlsx to HTML with base64 images using Aspose.Cells C# | Aspose.Cells C# convert workbook to HTML per worksheet | save Excel workbook as HTML with embedded pictures in .NET | C# code example for creating HTML representation of an Excel file using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export merged workbook to HTML | Aspose.Cells workbook.Save HTML format | generate HTML from Excel with embedded images | one page per sheet HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for a merged Excel file, loads it with Aspose.Cells, configures HtmlSaveOptions to embed images as Base64 (and optionally output one HTML page per worksheet), and saves the workbook as an HTML file while handling errors.
class Program
{
    static void Main()
    {
        const string inputPath = "mergedWorkbook.xlsx";
        const string outputPath = "mergedWorkbook.html";

        try
        {
            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the merged workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Optional: generate one HTML page per worksheet.
            // The OnePagePerSheet property may not be available in some versions; omit if unsupported.
            // htmlOptions.OnePagePerSheet = true;
            // Embed images directly into the HTML as Base64 strings
            htmlOptions.ExportImagesAsBase64 = true;

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML representation of the workbook has been created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
