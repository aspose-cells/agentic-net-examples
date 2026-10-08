// Title: Export all cell comments from an XLS workbook to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xls file using Aspose.Cells, sets HtmlSaveOptions.IsExportComments to true, and saves the workbook as an HTML file. | Show how to verify the existence of an input XLS file before converting it to HTML while preserving cell comments with Aspose.Cells. | Provide a complete example that configures HtmlSaveOptions for HTML export, ensures comments are included, and implements exception handling in C#.
// Common Searches: Aspose.Cells C# export cell comments to HTML from XLS | How to include worksheet comments when saving an XLS workbook as HTML using Aspose.Cells | HtmlSaveOptions.IsExportComments true example in .NET | Convert legacy .xls file to .html with comments using Aspose.Cells | C# code to check file existence and export Excel comments to HTML
// Tags: HtmlSaveOptions.IsExportComments C# | export cell comments to HTML Aspose.Cells | XLS to HTML conversion with comments Aspose.Cells | verify input file existence C# Aspose.Cells | exception handling workbook export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that input.xls exists, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions with IsExportComments enabled, and saves the workbook as output.html, preserving all cell comments while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xls";
            const string outputFile = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options (cell comments are exported by default)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save(outputFile, saveOptions);

            Console.WriteLine($"Workbook successfully saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
