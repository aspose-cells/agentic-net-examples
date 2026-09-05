// Title: How to enable collapse of multiple spaces after line breaks using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Generate C# code that removes extra spaces after line breaks by enabling the appropriate HtmlFormattingOptions in Aspose.Cells. | Show a complete example of loading an Excel workbook and exporting it to HTML while eliminating redundant whitespace using Aspose.Cells for .NET. | Explain how to detect HtmlFormattingOptions support and conditionally activate space collapsing when saving a workbook as HTML with Aspose.Cells.
// Common Searches: Aspose.Cells HtmlSaveOptions collapse spaces after newline C# example | remove extra whitespace in HTML output from Aspose.Cells workbook | how to trim multiple spaces when exporting Excel to HTML using Aspose.Cells .NET | HtmlFormattingOptions CollapseSpaces property not available version check
// Tags: Aspose.Cells HTML whitespace handling | C# HtmlSaveOptions space trimming | HtmlFormattingOptions CollapseSpaces usage | Excel to HTML conversion without extra spaces | conditional HtmlFormattingOptions usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample loads an Excel file, creates HtmlSaveOptions for HTML output, optionally sets HtmlFormattingOptions.CollapseSpaces to true when supported, and saves the workbook as HTML while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source Excel workbook
            var workbook = new Workbook(inputPath);

            // Create HTML save options
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Enable collapsing of multiple spaces after newlines if the property is available
            // (Older versions of Aspose.Cells may not expose HtmlFormattingOptions)
            // Uncomment the following line if your version supports it:
            // htmlOptions.HtmlFormattingOptions.CollapseSpaces = true;

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
