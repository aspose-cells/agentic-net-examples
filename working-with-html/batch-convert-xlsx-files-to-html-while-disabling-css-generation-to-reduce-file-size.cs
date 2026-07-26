// Title: C# Batch Convert XLSX to HTML without CSS using Aspose.Cells .NET
// Description: Scans a directory for *.xlsx files, creates matching *.html files, and converts each workbook with Aspose.Cells ConversionUtility. HtmlSaveOptions.DisableCss generates only inline styles, keeping the HTML lightweight.
// Keywords: Aspose.Cells | C# batch XLSX to HTML | DisableCss | HtmlSaveOptions | .NET Excel to HTML | convert multiple Excel files | reduce HTML size | ConversionUtility | inline style export | no external CSS
// Common Searches: aspocells batch convert xlsx to html c# | disable css when exporting Excel to HTML aspocells | convert folder of excel files to html .net | htmlsaveoptions disablecss example | aspocells conversionutility multiple files
// Developer Intent: Transform every .xlsx workbook in a folder into an .html file while suppressing external CSS generation.
// Use Cases: Publish a collection of spreadsheets as compact web pages for quick preview. | Embed Excel data in email bodies where external stylesheets are not allowed. | Automate nightly generation of intranet reports with minimal file size.
// AI Prompts: Generate C# code that converts a single XLSX file to HTML with Aspose.Cells, disabling CSS and adding custom inline formatting. | Show how to extend the batch conversion to wrap each HTML output in a custom template while keeping CSS disabled. | Explain how to parallelize the folder‑wide conversion using Aspose.Cells ConversionUtility to speed up processing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Scans a directory for *.xlsx files, creates matching *.html files, and converts each workbook with Aspose.Cells ConversionUtility. HtmlSaveOptions.DisableCss generates only inline styles, keeping the HTML lightweight.
class BatchXlsxToHtml
{
    static void Main()
    {
        // Folder containing source XLSX files
        string inputFolder = @"C:\InputXlsx";

        // Folder where HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all XLSX files in the input folder
        string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string sourcePath in xlsxFiles)
        {
            try
            {
                // Determine the destination HTML file path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Load options for XLSX format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // HTML save options with CSS generation disabled (inline styles only)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    DisableCss = true
                };

                // Perform the conversion using Aspose.Cells ConversionUtility
                ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                Console.WriteLine($"Converted: {sourcePath} -> {destPath} (CSS disabled)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file '{sourcePath}': {ex.Message}");
            }
        }
    }
}
