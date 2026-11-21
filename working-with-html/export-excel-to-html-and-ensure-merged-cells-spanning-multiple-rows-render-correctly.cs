// Title: Export an Excel worksheet to HTML with proper multi‑row merged cell rendering using Aspose.Cells for .NET
// Common Searches: how to keep merged cells spanning multiple rows when converting Excel to HTML with Aspose.Cells C# | Aspose.Cells C# export active worksheet to HTML preserving row‑spanning merges | C# HtmlSaveOptions example for UTF‑8 HTML output from an Excel workbook | save Excel workbook as HTML file using Aspose.Cells while handling missing input file | export Excel to HTML with merged cells not broken Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | C# export Excel to HTML with merged cells | Aspose.Cells multi‑row merged cell rendering HTML | UTF-8 HTML output from Excel using Aspose.Cells | Aspose.Cells workbook.Save HTML example

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The C# program loads an Excel file, verifies its existence, creates the output folder, and uses Aspose.Cells HtmlSaveOptions (ExportActiveWorksheetOnly = true, UTF‑8 encoding) to save the active worksheet as an HTML file. The export preserves merged cells that span multiple rows, ensuring correct rendering in the generated HTML.
class ExportExcelToHtml
{
    static void Main()
    {
        try
        {
            // Input Excel file path
            string excelPath = @"C:\Input\Sample.xlsx";

            // Verify that the source file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export only the active worksheet (set to false to export all sheets)
                ExportActiveWorksheetOnly = true,

                // Set the output encoding
                Encoding = Encoding.UTF8
            };

            // Output HTML file path
            string htmlPath = @"C:\Output\Sample.html";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(htmlPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML using the options
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Workbook successfully exported to HTML: {htmlPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
