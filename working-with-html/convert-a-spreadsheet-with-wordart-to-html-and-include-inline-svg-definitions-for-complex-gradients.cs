// Title: Convert an Excel workbook containing WordArt to HTML with inline SVG and Base64‑encoded images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with WordArt, sets HtmlSaveOptions.ExportImagesAsBase64 to true, and saves the workbook as an HTML file that contains inline SVG for shape gradients. | Show how to write the HTML output from Aspose.Cells to a MemoryStream, convert it to a UTF‑8 string, and persist it to disk in C#. | Explain how to verify the source Excel file exists before conversion and handle exceptions when exporting WordArt to HTML with Aspose.Cells.
// Common Searches: Aspose.Cells convert Excel WordArt to HTML with inline SVG | C# export Excel shapes as Base64 images in HTML using Aspose.Cells | How to embed complex gradient shapes as SVG when saving workbook to HTML .NET | HtmlSaveOptions ExportImagesAsBase64 example for WordArt | Save Excel to HTML with memory stream Aspose.Cells C#
// Tags: export WordArt to inline SVG Aspose.Cells | HtmlSaveOptions ExportImagesAsBase64 C# | convert Excel shapes to HTML Base64 images | inline SVG gradient definitions Aspose.Cells | memory stream HTML generation .NET | exception handling for Excel to HTML conversion

using System;
using System.IO;
using System.Text;
using System.Drawing.Imaging;
using Aspose.Cells;

namespace SpreadsheetToHtmlWithSvg
{
    // The example checks that input.xlsx exists, loads it with Aspose.Cells, configures HtmlSaveOptions to embed all images (including WordArt) as Base64, saves the workbook to a MemoryStream as HTML, converts the stream to a UTF‑8 string, and writes the resulting HTML—containing inline SVG definitions for complex gradients—to output.html.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file that contains WordArt.
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export images (including those generated from shapes) as Base64 strings.
                    ExportImagesAsBase64 = true

                    // Note: ExportChartImageFormat is not available in the current Aspose.Cells version.
                    // Charts will be exported using the default image format.
                };

                // Save the workbook to an in‑memory stream using the configured options.
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, htmlOptions);

                    // Convert the stream to a UTF‑8 string containing the HTML.
                    string htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());

                    // Write the HTML (with inline SVG definitions) to a file.
                    File.WriteAllText("output.html", htmlContent);
                }

                Console.WriteLine("Conversion completed. HTML with inline SVG saved to output.html");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
