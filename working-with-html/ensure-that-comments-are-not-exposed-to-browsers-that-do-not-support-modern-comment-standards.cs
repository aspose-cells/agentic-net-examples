// Title: Remove all worksheet comments and export an Excel workbook to HTML using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, clears every comment on each worksheet, and saves the workbook as an HTML file using HtmlSaveOptions. | Show a C# example that verifies the source Excel file exists, deletes all worksheet comments, and exports the workbook to HTML with basic error handling using Aspose.Cells.
// Common Searches: asp.net remove comments from Excel before exporting to HTML with Aspose.Cells | c# aspose.cells clear worksheet comments then save as html | how to hide Excel comments in HTML output using Aspose.Cells | export xlsx to html without comments asp.net
// Tags: Aspose.Cells clear worksheet comments C# | Aspose.Cells export workbook to HTML C# | remove Excel comments before HTML conversion Aspose | C# HtmlSaveOptions without comments Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the input Excel file, loads it with Aspose.Cells, removes all comments from every worksheet, and saves the workbook as an HTML file using default HtmlSaveOptions, while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the existing Excel file
            var workbook = new Workbook(inputPath);

            // Remove all comments so they are not rendered in the HTML output
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Comments.Clear();
            }

            // Configure HTML save options (default settings are sufficient)
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML without comments
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved to '{outputPath}' without comments.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
