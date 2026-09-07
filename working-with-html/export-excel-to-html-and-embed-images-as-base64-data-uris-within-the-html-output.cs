// Title: Convert an Excel workbook to HTML with embedded Base64 images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx file with Aspose.Cells, configures HtmlSaveOptions to embed all worksheet images as Base64 data URIs, and saves the result as an HTML file. | Show how to export only the active worksheet to HTML while embedding its pictures as Base64 using Aspose.Cells' HtmlSaveOptions in a .NET application. | Demonstrate retrieving the generated HTML as a UTF‑8 string from a MemoryStream after setting ExportImagesAsBase64, then writing it to disk with Aspose.Cells.
// Common Searches: how to export excel to html with embedded images using aspose.cells c# | aspnet convert xlsx to html base64 image data uri | aspose.cells save workbook as html string with base64 pictures | c# generate html from excel workbook with inline base64 images | aspose.cells HtmlSaveOptions ExportImagesAsBase64 example
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | export Excel to HTML with inline images | embed worksheet pictures as data URIs C# | save workbook to MemoryStream Aspose.Cells | convert xlsx to html string .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The sample verifies the input file, loads it into an Aspose.Cells Workbook, sets HtmlSaveOptions.ExportImagesAsBase64 to true, saves the workbook to a MemoryStream as HTML, reads the UTF‑8 HTML string, and writes it to an output file, handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Set HTML save options to embed images as Base64 data URIs
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = true   // Embed images directly in the HTML
                // ExportActiveWorksheetOnly = true // Optional: export only the active sheet
            };

            // Save the workbook to a memory stream to obtain the HTML as a string
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;
                string htmlContent = new StreamReader(ms, Encoding.UTF8).ReadToEnd();

                // Write the HTML output to a file (or use it as needed)
                File.WriteAllText(outputPath, htmlContent, Encoding.UTF8);
                Console.WriteLine($"Excel exported to HTML with embedded Base64 images: \"{outputPath}\"");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
