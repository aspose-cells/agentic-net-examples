// Title: C# batch conversion of multiple .xlsx workbooks to HTML with hidden sheets exported via Aspose.Cells
// AI Prompts: Generate a C# console program that iterates over a folder of .xlsx files and saves each workbook as an HTML file using Aspose.Cells, enabling HtmlSaveOptions.ExportHiddenWorksheet. | Write C# code that creates an output directory, loads each Excel workbook, configures hidden‑worksheet export, and logs success or error messages for every conversion. | Demonstrate exception handling and file‑existence checks when converting a collection of Excel workbooks to HTML with hidden sheets using Aspose.Cells.
// Common Searches: Aspose.Cells export hidden worksheets to HTML in C# batch job | Convert all .xlsx files in a directory to .html preserving hidden sheets using .NET | Example of HtmlSaveOptions.ExportHiddenWorksheet for multiple Excel files | C# console utility to transform Excel workbooks into HTML with hidden sheet data
// Tags: Aspose.Cells batch Excel to HTML conversion | HtmlSaveOptions hidden worksheet export C# | convert .xlsx files to .html with hidden sheets | C# console utility for Excel HTML export | process multiple workbooks using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // The C# console application scans a specified input folder for .xlsx files, ensures an output directory exists, loads each workbook with Aspose.Cells, sets HtmlSaveOptions.ExportHiddenWorksheet to true, saves the workbook as an HTML file, and logs the result or any errors for each file.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source Excel workbooks
            string sourceFolder = @"C:\InputExcels";

            // Folder where the generated HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (top level only)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Configure HTML save options to export hidden worksheets
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        ExportHiddenWorksheet = true
                    };

                    // Determine the output HTML file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
