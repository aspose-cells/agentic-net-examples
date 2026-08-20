// Title: C# Batch Convert Excel Workbooks to HTML with CSS Custom Properties using Aspose.Cells for .NET
// Description: A console utility that scans a folder for XLS, XLSX and XLSM files, creates an output directory, and converts each workbook to HTML with HtmlSaveOptions.EnableCssCustomProperties enabled. The option stores images as CSS variables, eliminating duplicate image data and reducing page size. Errors are logged per file.
// Keywords: Aspose.Cells batch conversion | Excel to HTML C# | EnableCssCustomProperties | CSS custom properties for images | deduplicate images Aspose | .NET Excel HTML export | ConversionUtility example | GitHub Aspose.Cells sample
// Common Searches: batch convert excel to html asp.net | enable css custom properties aspose.cells | remove duplicate images html export excel | c# convert folder of xlsx files to html | aspocells htmlsaveoptions css variables
// Developer Intent: Automatically transform multiple Excel files into HTML pages while using CSS custom properties to collapse repeated images into shared variables.
// Use Cases: Publish a library of financial spreadsheets as lightweight web reports with shared image assets. | Run a nightly job that converts newly uploaded Excel dashboards to static HTML for fast portal previews. | Create a CI pipeline that generates documentation from Excel specifications, minimizing bandwidth by deduplicating images.
// AI Prompts: Generate a C# script that logs each conversion result to a CSV file and sends a summary email after batch processing with Aspose.Cells. | Show how to modify the example to place each HTML file in a sub‑folder named after the source workbook and reference an external stylesheet. | Explain the mechanism behind EnableCssCustomProperties, how it creates CSS variables for images, and how to extract those variable definitions after conversion.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchExcelToHtml
{
    // A console utility that scans a folder for XLS, XLSX and XLSM files, creates an output directory, and converts each workbook to HTML with HtmlSaveOptions.EnableCssCustomProperties enabled. The option stores images as CSS variables, eliminating duplicate image data and reducing page size. Errors are logged per file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel files
            string inputFolder = @"C:\InputExcel";
            // Output folder for generated HTML files
            string outputFolder = @"C:\OutputHtml";

            // Ensure input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder '{inputFolder}' does not exist. Creating it.");
                Directory.CreateDirectory(inputFolder);
                Console.WriteLine("Place Excel files in the input folder and rerun the program.");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files (XLS, XLSX, XLSM) in the input folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string excelPath in excelFiles)
            {
                string extension = Path.GetExtension(excelPath).ToLowerInvariant();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                    continue; // Skip non‑Excel files

                // Determine output HTML file path
                string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                try
                {
                    // Create HtmlSaveOptions and enable CSS custom properties
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        EnableCssCustomProperties = true
                    };

                    // Use ConversionUtility with explicit LoadOptions and SaveOptions
                    LoadOptions loadOptions = new LoadOptions(); // default load options
                    ConversionUtility.Convert(excelPath, loadOptions, htmlPath, htmlOptions);

                    Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}' with CSS custom properties enabled.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
