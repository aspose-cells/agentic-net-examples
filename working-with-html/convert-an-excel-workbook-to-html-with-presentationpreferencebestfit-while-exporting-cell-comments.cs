// Title: Convert an Excel workbook to HTML and include cell comments using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel cell comments to HTML | how to save a workbook as HTML with comments using Aspose.Cells .NET | C# code example for converting .xlsx to HTML preserving comments | Aspose.Cells HtmlSaveOptions include cell comments in HTML output | convert Excel file to HTML with Aspose.Cells while keeping comments
// Tags: Aspose.Cells HtmlSaveOptions include comments | export Excel cell comments to HTML C# | convert workbook to HTML Aspose.Cells .NET | C# Aspose.Cells HTML conversion preserving comments | save Excel as HTML with cell annotations

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, creates default HtmlSaveOptions, and saves the workbook as output.html, handling any exceptions and reporting success or errors to the console.
class ExcelToHtmlConverter
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

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default options are sufficient)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Save the workbook as HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
