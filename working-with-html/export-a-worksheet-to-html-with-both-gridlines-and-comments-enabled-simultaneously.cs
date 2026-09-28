// Title: How to export an Excel worksheet to HTML with gridlines and cell comments using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file and saves it as an .html file with gridlines displayed using Aspose.Cells HtmlSaveOptions. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to turn on ExportGridLines and, when supported, ExportCellComments for HTML output. | Provide a complete example that checks for the source workbook, applies HTML export settings, and handles errors while preserving gridlines and comments.
// Common Searches: Aspose.Cells C# export worksheet to HTML with gridlines and comments | Enable cell comments in HTML output when saving Excel with Aspose.Cells | How to show gridlines in HTML files generated from Excel using Aspose.Cells | C# HtmlSaveOptions ExportGridLines and ExportCellComments example
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | Aspose.Cells HtmlSaveOptions ExportCellComments | C# Aspose.Cells export worksheet to HTML | HTML export with gridlines Aspose.Cells | cell comments HTML export Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program checks for the existence of input.xlsx, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions with ExportGridLines enabled (ExportCellComments is not available in this version), and saves the workbook as output.html while handling any exceptions.
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGridLines = true // Show gridlines in the HTML output
                // Note: ExportCellComments property is not available in this version of Aspose.Cells.
            };

            // Export the workbook to HTML
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
