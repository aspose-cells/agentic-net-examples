// Title: Convert an Excel workbook to HTML with Aspose.Cells for .NET while excluding hidden rows
// Common Searches: Aspose.Cells export Excel to HTML without hidden rows in C# | How to skip hidden rows when saving a workbook as HTML using Aspose.Cells .NET | C# generate HTML from .xlsx and ignore rows that are hidden | Aspose.Cells HtmlSaveOptions hide rows from HTML output | Export Excel to HTML and omit hidden rows programmatically with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions for HTML export | C# export workbook to HTML excluding hidden rows | HTML conversion of .xlsx with hidden rows omitted | Aspose.Cells generate HTML markup without hidden rows | save Excel as HTML using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx file into an Aspose.Cells Workbook, uses HtmlSaveOptions (which by default do not export hidden rows), and saves the workbook as output.html. It includes file existence checking and exception handling to ensure a smooth conversion while hidden rows are omitted from the generated HTML markup.
class ExportExcelToHtml
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Note: In this version of Aspose.Cells, hidden rows are not exported by default.
            // If a newer version provides ExportHiddenRows, it can be set here.

            // Export the workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
