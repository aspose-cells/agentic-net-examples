// Title: How to batch convert multiple Excel .xlsx workbooks to HTML with grid lines and comments using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates over every .xlsx file in a given folder and saves each workbook as an HTML file with grid lines and cell comments enabled via Aspose.Cells HtmlSaveOptions. | Create a reusable C# method that accepts an input Excel path and an output HTML path, sets HtmlSaveOptions.ExportGridLines = true and HtmlSaveOptions.IsExportComments = true, and exports the workbook to HTML.
// Common Searches: asp.net batch export excel files to html with grid lines and comments | c# Aspose.Cells convert folder of .xlsx to .html preserving comments | how to use HtmlSaveOptions ExportGridLines and IsExportComments in a loop | automate multiple workbook HTML conversion using Aspose.Cells C# | save each workbook as html with gridlines and comments Aspose.Cells example
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines C# | Aspose.Cells IsExportComments HTML export | batch convert Excel to HTML using Aspose.Cells | C# iterate over .xlsx files for HTML conversion | automated workbook HTML export with grid lines

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program scans a specified input directory for .xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file in an output directory. HtmlSaveOptions are configured with ExportGridLines and IsExportComments set to true, ensuring grid lines and cell comments appear in the generated HTML.
class HtmlExportBatch
{
    static void Main()
    {
        // Define the folder containing the source Excel files
        string sourceFolder = @"C:\InputWorkbooks";
        // Define the folder where the HTML files will be saved
        string outputFolder = @"C:\HtmlOutputs";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Gather all Excel files from the source folder (including subfolders if needed)
        List<string> excelFiles = new List<string>(Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly));

        // Process each workbook
        foreach (string excelPath in excelFiles)
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,      // Export grid lines
                IsExportComments = true      // Export cell comments
            };

            // Build the output HTML file name (same base name as the Excel file)
            string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlPath, saveOptions);

            // Release resources
            workbook.Dispose();
        }

        Console.WriteLine("Batch export completed.");
    }
}
