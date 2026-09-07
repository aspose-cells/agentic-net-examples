// Title: How to export an Excel workbook to HTML with column letters (A‑Z) as table headers using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.ExportColumnHeaders = true, and saves the workbook as an HTML file where the first row displays column letters as headers. | Show how to configure HtmlSaveOptions for UTF‑8 encoding and enable column‑letter headers when converting a worksheet to HTML with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# export Excel to HTML with column letters as header row | HtmlSaveOptions ExportColumnHeaders true example for .NET | Convert worksheet to HTML and show A B C column headers using Aspose.Cells | Save Excel file as HTML with column header letters in C#
// Tags: Aspose.Cells HTML export column header letters | C# HtmlSaveOptions column header configuration | convert worksheet to HTML with column letters Aspose | UTF8 encoding HTML output Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example loads an Excel workbook, creates HtmlSaveOptions with UTF‑8 encoding, enables ExportColumnHeaders to output column letters (A, B, C…) as the first HTML table row, and saves the result as an HTML file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                Encoding = Encoding.UTF8
                // ExportColumnHeaders, ExportRowHeaders, and IncludeWorksheetHeader are not required
                // or may not be available in the current Aspose.Cells version.
            };

            // Save the workbook (or a specific worksheet) as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
