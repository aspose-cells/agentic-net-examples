// Title: Export defined print area of an Excel workbook to HTML using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.ExportPrintAreaOnly to true, and saves the result as an HTML file. | Write error‑handling logic that verifies the source Excel file exists before converting it to HTML with a print‑area‑only export. | Show how to customize additional HtmlSaveOptions (such as page title or encoding) while exporting only the defined print area to HTML using Aspose.Cells.
// Common Searches: Aspose.Cells C# convert Excel workbook to HTML limited to the set print area | How to enable ExportPrintAreaOnly in HtmlSaveOptions when saving as HTML in .NET | Sample code for saving an .xlsx file as HTML with only the print region using Aspose.Cells | C# Aspose.Cells HTML export with print area restriction example | Using HtmlSaveOptions to generate HTML from Excel with print area constraints
// Tags: Aspose.Cells HtmlSaveOptions ExportPrintAreaOnly | C# export Excel print area to HTML | Aspose.Cells HTML conversion with print area restriction | save workbook as HTML using Aspose.Cells | print area only HTML output Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook, configures HtmlSaveOptions with ExportPrintAreaOnly = true, and saves the workbook as an HTML file containing only the defined print area, including basic file‑existence error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to export only the defined print area
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportPrintAreaOnly = true
            };

            // Save the workbook as HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
