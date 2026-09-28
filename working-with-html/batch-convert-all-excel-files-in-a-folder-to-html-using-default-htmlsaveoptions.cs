// Title: C# console app to batch convert Excel workbooks (.xls, .xlsx, .xlsm, .xlsb) in a folder to HTML using Aspose.Cells default HtmlSaveOptions
// AI Prompts: Generate a C# console program that scans a given directory for .xls, .xlsx, .xlsm, and .xlsb files, creates an "HtmlOutput" subfolder, and saves each workbook as an HTML file with Aspose.Cells SaveFormat.Html using the default HtmlSaveOptions. | Write C# code that enumerates all Excel files in a folder (non‑recursive), loads each with Aspose.Cells Workbook, and exports them to HTML files in a separate output directory while logging the conversion results. | Provide a C# script that uses Aspose.Cells to batch convert multiple Excel workbooks to HTML, handling folder creation, file naming, and default HTML save settings.
// Common Searches: how to use Aspose.Cells in C# to convert a folder of Excel files to HTML | C# batch conversion of .xls and .xlsx files to HTML with Aspose.Cells default options | example code for converting multiple Excel workbooks to HTML using Aspose.Cells SaveFormat.Html | create output subfolder and export Excel to HTML in a C# console application
// Tags: Aspose.Cells batch Excel to HTML conversion | C# convert Excel files to HTML using SaveFormat.Html | default HtmlSaveOptions Aspose.Cells | enumerate Excel files in folder C# Aspose.Cells | create HtmlOutput subfolder C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// A C# console application that scans a specified directory for .xls, .xlsx, .xlsm, and .xlsb files, creates an "HtmlOutput" subfolder, loads each workbook with Aspose.Cells, and saves it as an HTML file using the default HtmlSaveOptions (SaveFormat.Html).
class ExcelToHtmlBatchConverter
{
    static void Main(string[] args)
    {
        // Specify the folder containing Excel files.
        // You can change this path as needed.
        string sourceFolder = @"C:\ExcelFiles";

        // Specify the folder where HTML files will be saved.
        // If the folder does not exist, it will be created.
        string outputFolder = Path.Combine(sourceFolder, "HtmlOutput");
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Define the Excel file extensions to process.
        string[] excelExtensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

        // Get all Excel files in the source folder (non‑recursive).
        var excelFiles = Directory.GetFiles(sourceFolder)
                                  .Where(f => excelExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

        foreach (var excelPath in excelFiles)
        {
            // Load the workbook from the Excel file.
            Workbook workbook = new Workbook(excelPath);

            // Prepare the output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the workbook as HTML using default HtmlSaveOptions.
            workbook.Save(htmlPath, SaveFormat.Html);
            
            Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
