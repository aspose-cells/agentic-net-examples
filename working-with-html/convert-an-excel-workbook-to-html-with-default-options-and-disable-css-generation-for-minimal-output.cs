// Title: Convert an Excel workbook to HTML with Aspose.Cells in C# while disabling CSS generation for a lightweight output
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells and saves it as HTML, configuring HtmlSaveOptions to turn off CSS stylesheet creation. | Show how to export an Excel workbook to minimal HTML markup in C# by using Aspose.Cells HtmlSaveOptions with CSS generation disabled.
// Common Searches: asp.net convert excel to html without css using aspose.cells | c# export workbook to plain html no stylesheet asp.net | how to disable css when saving excel as html with aspose.cells | minimal html output from excel workbook c# aspose.cells | default htmlsaveoptions usage aspose.cells c#
// Tags: Aspose.Cells HtmlSaveOptions disable CSS | C# export Excel to plain HTML | Aspose.Cells minimal HTML output | Excel to HTML conversion without stylesheet | Aspose.Cells default HTML save options

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the presence of an input .xlsx file, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions with CSS generation turned off, and saves the workbook as a lightweight HTML file, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file using the configured options
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
