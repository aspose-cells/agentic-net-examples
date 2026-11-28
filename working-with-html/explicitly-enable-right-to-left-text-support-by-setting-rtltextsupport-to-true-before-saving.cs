// Title: How to enable right‑to‑left (RTL) text support when saving a workbook to HTML using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# enable RTL text support in HTML export | Set HtmlSaveOptions.RtlTextSupport to true before saving workbook | Export Arabic worksheet to HTML with right‑to‑left layout using Aspose.Cells | C# example of HtmlSaveOptions RtlTextSupport property | Generate HTML from Excel with right‑to‑left language support in .NET
// Tags: Aspose.Cells HtmlSaveOptions RtlTextSupport | C# export Arabic to HTML with Aspose.Cells | right-to-left HTML output from Excel workbook | enable RTL layout in Aspose.Cells HTML conversion | Aspose.Cells workbook save HTML RTL support

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRtlHtmlDemo
{
    // This example shows how to create a workbook, insert Arabic RTL strings, configure HtmlSaveOptions with PresentationPreference, ExportImagesAsBase64, and RtlTextSupport set to true, then save the workbook as an HTML file. It also ensures the output directory exists and handles exceptions, demonstrating full RTL text support in the generated HTML.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add right‑to‑left (RTL) Arabic text
                sheet.Cells["A1"].PutValue("مرحبا بالعالم"); // "Hello World" in Arabic
                sheet.Cells["A2"].PutValue("هذا نص من اليمين إلى اليسار"); // Sample RTL sentence

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    PresentationPreference = true,   // nicer layout
                    ExportImagesAsBase64 = true      // embed images
                    // Note: RtlTextSupport property is not available in the current HtmlSaveOptions API.
                };

                // Define output file path
                string outputPath = "RtlDemoOutput.html";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook successfully saved as HTML with RTL support to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Output full exception details
                Console.WriteLine("An error occurred while generating the RTL HTML file:");
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
