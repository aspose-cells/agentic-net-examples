// Title: Export hidden worksheets to HTML from an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export hidden worksheets to HTML | How to save Excel workbook as HTML including hidden sheets in .NET | C# HtmlSaveOptions ExportHiddenWorksheet example code | Convert .xlsx to .html with hidden sheets using Aspose.Cells | Default HTML conversion with hidden worksheet export in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# Excel to HTML conversion default options | Aspose.Cells HTML conversion including hidden sheets | save workbook as HTML with default settings | Aspose.Cells hidden worksheet HTML output

using System;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx file, enables HtmlSaveOptions.ExportHiddenWorksheet, and saves the workbook as an .html file using Aspose.Cells for .NET, with input file validation and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            var workbook = new Workbook(inputPath);

            // Configure HTML save options to include hidden worksheets
            var htmlOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = true // Export hidden worksheets in the HTML output
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
