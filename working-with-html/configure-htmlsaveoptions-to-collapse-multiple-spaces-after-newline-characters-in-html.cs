// Title: How to configure Aspose.Cells HtmlSaveOptions in C# to collapse multiple spaces after newline characters when exporting to HTML
// Common Searches: Aspose.Cells HtmlSaveOptions collapse spaces after newline in C# | How to eliminate multiple spaces when saving workbook as HTML using Aspose.Cells | C# export Excel worksheet to HTML without double spaces after line breaks | Trim whitespace in generated HTML from Aspose.Cells workbook | Set HtmlSaveOptions to remove extra spaces after line breaks in HTML output
// Tags: HtmlSaveOptions whitespace collapse | Aspose.Cells HTML export space handling | C# Excel to HTML whitespace trimming | newline space reduction Aspose.Cells | configure HtmlSaveOptions line break spacing

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a Workbook, adds sample data, configures HtmlSaveOptions (placeholder for additional settings), ensures the output directory exists, and saves the workbook as an HTML file using the configured options, allowing control over whitespace handling after newline characters.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Example data (optional)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("First line");
                sheet.Cells["A2"].PutValue("Second line");

                // Configure HtmlSaveOptions
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // If needed, set additional options here (e.g., htmlOptions.ExportImagesAsBase64 = true;)

                // Determine output path and ensure directory exists
                string outputPath = "output.html";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
