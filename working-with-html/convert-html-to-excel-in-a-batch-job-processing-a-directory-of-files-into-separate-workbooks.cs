using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input directory containing HTML files
            string inputDirectory = @"C:\InputHtml";
            // Output directory where converted Excel files will be saved
            string outputDirectory = @"C:\OutputExcel";

            try
            {
                // Verify input directory exists
                if (!Directory.Exists(inputDirectory))
                {
                    Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                    return;
                }

                // Ensure the output directory exists
                if (!Directory.Exists(outputDirectory))
                    Directory.CreateDirectory(outputDirectory);

                // Get all HTML files in the input directory (non‑recursive)
                string[] htmlFiles = Directory.GetFiles(inputDirectory, "*.html");

                foreach (string htmlFilePath in htmlFiles)
                {
                    try
                    {
                        // Verify the HTML file exists before loading
                        if (!File.Exists(htmlFilePath))
                        {
                            Console.WriteLine($"File not found: {htmlFilePath}");
                            continue;
                        }

                        // Prepare the output Excel file path (same name, .xlsx extension)
                        string excelFileName = Path.GetFileNameWithoutExtension(htmlFilePath) + ".xlsx";
                        string excelFilePath = Path.Combine(outputDirectory, excelFileName);

                        // Load the HTML file into a workbook using LoadOptions for HTML format
                        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

                        // Save the workbook as an Excel file (XLSX format)
                        workbook.Save(excelFilePath, SaveFormat.Xlsx);

                        Console.WriteLine($"Converted: {Path.GetFileName(htmlFilePath)} -> {excelFileName}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting '{htmlFilePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}