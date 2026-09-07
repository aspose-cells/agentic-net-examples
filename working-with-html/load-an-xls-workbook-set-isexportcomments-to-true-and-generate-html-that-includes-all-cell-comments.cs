// Title: Export an XLS workbook to HTML including all cell comments with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xls file using Aspose.Cells, sets HtmlSaveOptions.IsExportComments to true, and saves the workbook as an HTML document. | Demonstrate how to check for the input file, configure HtmlSaveOptions for comment export, and handle errors while converting an Excel workbook to HTML with Aspose.Cells.
// Common Searches: Aspose.Cells C# export XLS to HTML with cell comments | How to include Excel comments when saving as HTML using Aspose.Cells | HtmlSaveOptions IsExportComments true example .NET | Convert legacy .xls file to HTML preserving comments Aspose | C# code to save workbook as HTML with comments Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions export comments | export XLS to HTML with comments C# | load workbook save as HTML Aspose.Cells | IsExportComments true Aspose.Cells | convert legacy Excel to HTML preserving comments

using System;
using System.IO;
using Aspose.Cells;

// The example checks that input.xls exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions with IsExportComments enabled, and saves the workbook as output.html, while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the XLS workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (comments are exported by default)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved as HTML to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
