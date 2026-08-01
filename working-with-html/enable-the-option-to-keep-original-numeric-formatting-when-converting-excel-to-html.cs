// Title: Preserve Numeric Formatting When Exporting Excel to HTML with Aspose.Cells for .NET
// Description: Shows how to enable Aspose.Cells' HtmlSaveOptions.ExportNumericFormat property in C# so that number, currency, percentage, and custom formats from an Excel workbook are retained in the generated HTML.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportNumericFormat | Excel to HTML conversion | preserve numeric format | C# Excel HTML export | retain number formatting | currency format HTML | percentage format export | custom number format Aspose
// Common Searches: Aspose.Cells ExportNumericFormat example | keep Excel number formatting in HTML | C# export Excel to HTML with original numeric styles | HtmlSaveOptions preserve currency symbols | how to retain custom number formats when converting Excel to HTML
// Developer Intent: Enable the ExportNumericFormat flag in HtmlSaveOptions so the HTML output mirrors the workbook’s numeric styling.
// Use Cases: Financial reports on a web portal must display exact currency symbols and decimal precision. | Dashboard widgets that show percentages need to match Excel source formatting. | Web‑based spreadsheet viewer that presents custom number formats without visual loss. | Automated documentation generation where numeric precision is critical.
// AI Prompts: Write C# code using Aspose.Cells to convert an .xlsx file to HTML with ExportNumericFormat enabled. | Explain additional HtmlSaveOptions settings that affect numeric display, such as PreserveFormatting and ExportDateTimeFormat. | Provide a step‑by‑step test plan to verify that numeric formats remain unchanged after HTML export.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to enable Aspose.Cells' HtmlSaveOptions.ExportNumericFormat property in C# so that number, currency, percentage, and custom formats from an Excel workbook are retained in the generated HTML.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.html";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options (default settings preserve formatting)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save as HTML
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
