// Title: Export hidden worksheets to HTML with Aspose.Cells for .NET – guidance for large workbook performance
// AI Prompts: Generate C# code that loads an Excel workbook and saves it as an HTML file while preserving hidden worksheets using Aspose.Cells HtmlSaveOptions. | Write a C# example that configures HtmlSaveOptions to include hidden sheets in the HTML output and demonstrates saving a large workbook efficiently with Aspose.Cells.
// Common Searches: asp.net aspose.cells export hidden worksheets to html | c# htmlsaveoptions include hidden sheets aspose.cells | optimizing html export performance for large Excel files using aspose.cells | how to convert Excel to html and keep hidden worksheets in .net | aspose.cells html conversion settings for big workbooks
// Tags: Aspose.Cells HtmlSaveOptions export hidden worksheets | C# convert Excel to HTML performance | large workbook HTML export Aspose.Cells | include hidden sheets in HTML output Aspose.Cells | Aspose.Cells HTML conversion settings | optimize Aspose.Cells HTML export for big files

using System;
using System.IO;
using Aspose.Cells;

// // Loads an Excel file, sets HtmlSaveOptions.ExportHiddenWorksheet = true, and saves the workbook as an HTML document using Aspose.Cells for .NET.
class ExportHtmlExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export hidden worksheets as part of the HTML output
                ExportHiddenWorksheet = true
                // HtmlCrossType property is not available in this version of Aspose.Cells
            };

            // Save the workbook as an HTML file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
