// Title: Keep Excel worksheet name capitalization in HTML heading tags when exporting with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export Excel to HTML preserving worksheet name case | How to retain original sheet name capitalization in generated HTML using Aspose.Cells | Aspose.Cells HtmlSaveOptions keep sheet title case in HTML output | Export workbook to HTML without changing heading tag case in .NET | Preserve Excel tab name case when converting to HTML with Aspose.Cells
// Tags: Aspose.Cells HTML export retain sheet name case | C# HtmlSaveOptions preserve sheet title capitalization | Excel to HTML heading tag case handling | Workbook.Save as HTML case-sensitive sheet names | maintain original worksheet case in generated HTML

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, uses the default HtmlSaveOptions (which keep the original worksheet name's capitalization), and saves the file as HTML, ensuring that the generated heading tags reflect the exact case of the source sheet names.
class PreserveWorksheetNameCase
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set HTML save options (default behavior preserves worksheet name case)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
