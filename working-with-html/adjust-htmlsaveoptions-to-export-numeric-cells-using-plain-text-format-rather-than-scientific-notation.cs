// Title: Export numeric cells as plain text (no scientific notation) in HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to HTML without scientific notation | how to force plain number format when exporting Excel to HTML using Aspose.Cells | set custom number format for cells before HTML export in Aspose.Cells .NET | prevent scientific notation in HTML output from Aspose.Cells workbook | HtmlSaveOptions to keep numbers as plain text in Aspose.Cells
// Tags: HTML export plain numeric format Aspose.Cells | disable scientific notation HtmlSaveOptions | custom cell number format Aspose.Cells C# | non‑scientific number rendering HTML Aspose.Cells | Aspose.Cells numeric formatting for HTML output

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, applies custom number formats ("0" for integers and "0.##############" for decimals) to cells, configures HtmlSaveOptions, and saves the file as HTML so that numeric values appear as plain text instead of scientific notation.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Put some numeric values that would normally be shown in scientific notation
                sheet.Cells["A1"].PutValue(123456789);
                sheet.Cells["A2"].PutValue(0.00000123);

                // Apply a custom number format that forces plain (non‑scientific) representation
                // For large integers use "0", for small decimals use a format with enough placeholders
                var styleA1 = sheet.Cells["A1"].GetStyle();
                styleA1.Custom = "0";
                sheet.Cells["A1"].SetStyle(styleA1);

                var styleA2 = sheet.Cells["A2"].GetStyle();
                styleA2.Custom = "0.##############";
                sheet.Cells["A2"].SetStyle(styleA2);

                // Configure HTML save options (default options are sufficient for plain text)
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Ensure the output directory exists
                string outputPath = "ExportPlainText.html";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to HTML using the configured options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
