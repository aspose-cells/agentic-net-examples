// Title: Batch convert multiple .xlsx workbooks to HTML while preserving conditional formatting with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a directory for .xlsx files and saves each workbook as an HTML file using Aspose.Cells, ensuring conditional formatting is exported. | Show how to configure Aspose.Cells HTML save options to keep conditional formatting during a bulk conversion of Excel workbooks to HTML. | Provide C# code that creates an output folder, loads each Excel file from an input folder, and calls Workbook.Save with appropriate options to generate HTML files that retain visual rules.
// Common Searches: aspocells convert a folder of xlsx files to html with conditional formatting c# | c# loop through directory convert excel workbooks to html using Aspose.Cells | how to export conditional formatting to html with Aspose.Cells SaveOptions | convert multiple Excel files to html programmatically Aspose.Cells .NET | html save options preserve visual rules Aspose.Cells conversion
// Tags: batch xlsx to html Aspose.Cells | export conditional formatting html Aspose.Cells | c# iterate directory convert excel to html | html save options Aspose.Cells example | preserve visual rules Aspose.Cells html conversion

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // A C# console application that enumerates all .xlsx files in a specified input folder, loads each workbook with Aspose.Cells, and saves it as an HTML file in an output folder using HTML save options that export conditional formatting, thereby retaining the original visual rules.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel workbooks
            string inputFolder = @"C:\InputExcelFiles";

            // Output folder for generated HTML files
            string outputFolder = @"C:\OutputHtmlFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files (adjust pattern as needed)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            // Configure HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    // Verify the Excel file still exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found, skipping: {excelPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Determine output HTML file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML with the specified options
                    workbook.Save(htmlPath, htmlOptions);
                    Console.WriteLine($"Converted: {excelPath} -> {htmlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
