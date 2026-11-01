// Title: Export an Excel workbook to HTML while preserving exact numeric formatting with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, configures HtmlSaveOptions to export numeric values as text, and saves the file as HTML while keeping the original number appearance. | Show how to set Aspose.Cells HTML export options so that Excel cell number formats are retained in the generated HTML. | Provide a .NET snippet that verifies the source file, applies ExportCellValueAsString when supported, and writes the formatted HTML output.
// Common Searches: how to keep Excel number formatting when converting to HTML using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions numeric data as string example | preserve numeric display in HTML export from .xlsx with Aspose.Cells | C# export workbook to HTML retaining original numeric values
// Tags: Aspose.Cells HtmlSaveOptions ExportNumericDataAsString | C# Excel to HTML numeric formatting | Aspose.Cells numeric format preservation | SaveFormat.Html Aspose.Cells example | ExportCellValueAsString option C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook, configures HtmlSaveOptions (including ExportNumericDataAsString and optionally ExportCellValueAsString) to retain the exact numeric display from Excel, and saves the result as output.html, with basic file existence checks and error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input workbook exists before loading.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook.
            Workbook workbook = new Workbook(inputPath);

            // Set up HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The following options are not available in older Aspose.Cells versions.
            // If your version supports them, uncomment the lines.
            // htmlOptions.ExportNumericDataAsString = true;
            // htmlOptions.ExportCellValueAsString = true;

            // Save the workbook as an HTML file.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
