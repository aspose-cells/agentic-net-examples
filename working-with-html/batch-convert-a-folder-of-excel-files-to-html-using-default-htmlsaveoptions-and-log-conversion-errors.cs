// Title: Convert a folder of Excel workbooks to HTML with Aspose.Cells for .NET and log conversion errors
// AI Prompts: Write a C# console application that scans a specified directory for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and saves it as an .html file using the default HtmlSaveOptions. | Enhance the program to accept input and output folder paths via command‑line arguments and automatically create the output directory if it does not exist. | Add robust error handling that catches any exception during loading or saving and writes the offending file name and exception message to the error console or a log file.
// Common Searches: aspocells c# batch convert excel files to html with error logging | how to process multiple xlsx files and save as html using Aspose.Cells | c# console app convert folder of xls to html and capture conversion failures | default HtmlSaveOptions usage in Aspose.Cells for bulk conversion
// Tags: Aspose.Cells bulk Excel to HTML conversion | C# folder enumeration for .xls .xlsx .xlsm files | HtmlSaveOptions default settings Aspose.Cells | error handling Aspose.Cells workbook conversion | command line input output paths C# console

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that iterates over all .xls, .xlsx, and .xlsm files in a given input folder, converts each workbook to an HTML file using Aspose.Cells' default HtmlSaveOptions, saves the results to a specified output folder (creating it if needed), and logs any conversion errors to the error stream.
class ExcelToHtmlBatchConverter
{
    static void Main(string[] args)
    {
        // Input folder containing Excel files (default if not provided)
        string inputFolder = args.Length > 0 ? args[0] : @"C:\InputExcel";

        // Output folder for generated HTML files (default if not provided)
        string outputFolder = args.Length > 1 ? args[1] : @"C:\OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all files in the input folder
        string[] allFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in allFiles)
        {
            // Process only Excel file extensions
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                continue;

            try
            {
                // Load the Excel workbook
                Workbook workbook = new Workbook(filePath);

                // Determine the HTML output path
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML using default HtmlSaveOptions
                workbook.Save(htmlPath, SaveFormat.Html);

                Console.WriteLine($"Converted: {filePath} -> {htmlPath}");
            }
            catch (Exception ex)
            {
                // Log any conversion errors
                Console.Error.WriteLine($"Error converting '{filePath}': {ex.Message}");
            }
        }
    }
}
