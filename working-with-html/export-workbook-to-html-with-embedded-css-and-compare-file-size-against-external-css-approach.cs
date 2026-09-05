// Title: Generate HTML from an Aspose.Cells workbook with embedded CSS and compare its size to the external CSS version using C#
// AI Prompts: Write C# code that saves a Workbook as HTML using the default external stylesheet, then reads the generated style.css and inserts it into the <head> to create a single HTML file with embedded CSS. | Add logic to measure the byte size of the HTML file that contains the embedded CSS and the combined size of the separate HTML and CSS files, then output which method results in a smaller total size. | Configure HtmlSaveOptions to export only the active worksheet when converting the workbook to HTML.
// Common Searches: Aspose.Cells C# generate HTML with CSS inside the file | how to embed generated style.css into Aspose.Cells HTML output | compare size of HTML with embedded CSS versus external CSS in Aspose.Cells | C# calculate total output size of Aspose.Cells HTML export with stylesheet
// Tags: Aspose.Cells HTML export with embedded stylesheet | C# embed external CSS into generated HTML file | HtmlSaveOptions ExportActiveWorksheetOnly example | compare embedded CSS vs external CSS file size | calculate HTML and CSS output size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to export a workbook to HTML using Aspose.Cells, create both an external CSS file and an HTML file with the CSS embedded, and then compare their byte sizes to determine which approach yields a smaller overall output.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                // Populate the worksheet with data
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define output folder
                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
                Directory.CreateDirectory(outputDir);

                // -------------------------------------------------
                // Export HTML with external CSS (default behavior)
                // -------------------------------------------------
                string externalHtmlPath = Path.Combine(outputDir, "ExternalCss.html");
                HtmlSaveOptions externalOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = true
                };
                workbook.Save(externalHtmlPath, externalOptions);

                // Determine the generated CSS file name (default is "style.css")
                string cssFileName = "style.css";
                string cssFilePath = Path.Combine(outputDir, cssFileName);

                // -------------------------------------------------
                // Create HTML with embedded CSS by inlining the CSS file
                // -------------------------------------------------
                string embeddedHtmlPath = Path.Combine(outputDir, "EmbeddedCss.html");
                if (File.Exists(cssFilePath))
                {
                    string cssContent = File.ReadAllText(cssFilePath);
                    string htmlContent = File.ReadAllText(externalHtmlPath);

                    // Insert CSS into <head> section
                    int headCloseIndex = htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
                    if (headCloseIndex >= 0)
                    {
                        string styleTag = $"<style type=\"text/css\">\n{cssContent}\n</style>\n";
                        htmlContent = htmlContent.Insert(headCloseIndex, styleTag);
                    }
                    else
                    {
                        // Fallback: prepend CSS at the beginning of the file
                        htmlContent = $"<style type=\"text/css\">\n{cssContent}\n</style>\n{htmlContent}";
                    }

                    File.WriteAllText(embeddedHtmlPath, htmlContent);
                }
                else
                {
                    // If CSS file was not generated, just copy the external HTML as embedded version
                    File.Copy(externalHtmlPath, embeddedHtmlPath, true);
                }

                // -------------------------------------------------
                // Compare file sizes
                // -------------------------------------------------
                long embeddedSize = new FileInfo(embeddedHtmlPath).Length;
                long externalHtmlSize = new FileInfo(externalHtmlPath).Length;
                long externalCssSize = File.Exists(cssFilePath) ? new FileInfo(cssFilePath).Length : 0;
                long totalExternalSize = externalHtmlSize + externalCssSize;

                Console.WriteLine($"Embedded CSS HTML size: {embeddedSize} bytes");
                Console.WriteLine($"External CSS HTML size: {externalHtmlSize} bytes");
                Console.WriteLine($"External CSS file size: {externalCssSize} bytes");
                Console.WriteLine($"Total size (HTML + CSS): {totalExternalSize} bytes");

                if (embeddedSize < totalExternalSize)
                {
                    Console.WriteLine("Embedded CSS approach results in a smaller overall file size.");
                }
                else if (embeddedSize > totalExternalSize)
                {
                    Console.WriteLine("External CSS approach results in a smaller overall file size.");
                }
                else
                {
                    Console.WriteLine("Both approaches produce the same total file size.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
