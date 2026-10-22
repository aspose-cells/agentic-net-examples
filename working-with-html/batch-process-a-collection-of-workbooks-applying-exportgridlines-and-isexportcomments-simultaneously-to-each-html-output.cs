// Title: Batch convert a folder of Excel .xlsx workbooks to HTML with grid lines and cell comments using Aspose.Cells for .NET
// AI Prompts: Generate C# code that scans a specified directory for .xlsx files, loads each workbook with Aspose.Cells, sets HtmlSaveOptions.ExportGridLines and HtmlSaveOptions.IsExportComments to true, and saves the result as .html files in an output folder. | Write a .NET console application that iterates over all Excel workbooks in a folder and exports each to HTML while preserving grid lines and cell comments via Aspose.Cells HtmlSaveOptions.
// Common Searches: asp.net batch convert excel files to html with grid lines and comments using Aspose.Cells | c# Aspose.Cells export multiple .xlsx workbooks to .html preserving cell comments | how to use HtmlSaveOptions ExportGridLines for bulk Excel to HTML conversion | automate folder-wide Excel to HTML conversion with Aspose.Cells in C#
// Tags: batch html export Aspose.Cells | HtmlSaveOptions ExportGridLines C# | Aspose.Cells export cell comments to HTML | convert multiple .xlsx to .html C# | C# folder iteration Aspose.Cells conversion

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample scans an input directory for .xlsx files, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to enable ExportGridLines and IsExportComments, and saves each workbook as an .html file in a designated output folder.
class HtmlExportBatch
{
    static void Main()
    {
        // Define the folder containing the source Excel files
        string sourceFolder = @"C:\InputWorkbooks";
        // Define the folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

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
        }

        Console.WriteLine("Batch export completed.");
    }
}
