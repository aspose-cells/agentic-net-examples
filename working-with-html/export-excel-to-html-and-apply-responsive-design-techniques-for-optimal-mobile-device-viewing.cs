// Title: Convert an Excel workbook to mobile‑friendly HTML with embedded Base64 images using Aspose.Cells for .NET (C#)
// Common Searches: C# how to export an .xlsx file to HTML with images embedded as Base64 using Aspose.Cells | Aspose.Cells generate responsive HTML for Excel workbook in .NET | Save Excel workbook as HTML for mobile devices with Aspose.Cells C# example | Embedding images as Base64 when converting Excel to HTML with Aspose.Cells | Create mobile‑optimized HTML from Excel using Aspose.Cells HtmlSaveOptions
// Tags: Aspose.Cells HtmlSaveOptions embed images base64 | C# export Excel to mobile-friendly HTML | Aspose.Cells convert .xlsx to responsive HTML | HTML export options for Excel workbook .NET | Base64 image embedding in Aspose.Cells HTML output

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook), configures HtmlSaveOptions to embed images as Base64 strings, and saves the workbook as an HTML file. While Aspose.Cells does not provide a direct responsive‑layout flag, the generated HTML can be styled for mobile devices after export.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Optionally add some sample data to avoid an empty workbook.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");
            }

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = true // Embed images directly in the HTML.
                // Note: Responsive layout is not directly supported via a property in this version of Aspose.Cells.
            };

            // Save the workbook as HTML.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
