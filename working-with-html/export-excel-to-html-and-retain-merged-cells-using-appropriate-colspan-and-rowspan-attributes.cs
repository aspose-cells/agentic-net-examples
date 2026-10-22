// Title: Export an Excel workbook to HTML with merged cells preserved using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as HTML, ensuring merged cells are rendered with correct colspan and rowspan attributes. | Demonstrate how to configure HtmlSaveOptions in Aspose.Cells so that merged ranges remain intact when converting a workbook to HTML.
// Common Searches: how to export merged cells from Excel to HTML using Aspose.Cells .NET | C# Aspose.Cells HtmlSaveOptions preserve merged ranges | generate HTML from .xlsx with correct colspan and rowspan attributes | Aspose.Cells export workbook to HTML while keeping cell merges | save Excel file as HTML with merged cells using Aspose.Cells for .NET
// Tags: Aspose.Cells HtmlSaveOptions merged cells | export Excel to HTML with colspan rowspan | C# Aspose.Cells preserve merged ranges | HTML conversion of .xlsx using Aspose | save workbook as HTML Aspose .NET

using System;
using System.IO;
using Aspose.Cells;

// Loads an .xlsx workbook with Aspose.Cells, applies HtmlSaveOptions (UTF‑8 encoding), and saves it as HTML; merged cells are automatically emitted with appropriate colspan and rowspan attributes.
class ExportExcelToHtml
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Set the encoding for the generated HTML
                Encoding = System.Text.Encoding.UTF8
                // ExportMergedCells property is not required; merged cells are handled by default
            };

            // Save the workbook as an HTML file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully exported to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
