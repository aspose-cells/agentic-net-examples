// Title: Batch convert multiple Excel workbooks to PDF with individual PdfSaveOptions using Aspose.Cells for .NET
// AI Prompts: Write C# code that loops through a collection of Excel file paths, loads each workbook with Aspose.Cells, and saves it as a PDF using a separate PdfSaveOptions instance for every file. | Demonstrate how to apply the OnePagePerSheet option only to selected Excel files while performing a batch Excel‑to‑PDF conversion, including try‑catch error handling for missing source files. | Create a robust routine that checks whether each source workbook exists, creates the target output folder if necessary, and logs success or failure for each PDF conversion with custom options.
// Common Searches: how to batch convert excel to pdf with different PdfSaveOptions in C# Aspose.Cells | Aspose.Cells convert multiple .xlsx files to PDF each with its own settings | C# example for per‑file OnePagePerSheet option during Excel to PDF conversion | error handling for missing Excel files in batch PDF conversion using Aspose.Cells | automatically create output folder when saving PDF from Excel with Aspose.Cells
// Tags: batch excel to pdf conversion Aspose.Cells | per‑file PdfSaveOptions configuration | OnePagePerSheet PDF setting | file existence validation Aspose.Cells | automatic output directory creation C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToPdf
{
    // Holds conversion information for each Excel file
    // The program defines a ConversionTask class that stores the source Excel path, destination PDF path, and a PdfSaveOptions object. It builds a list of such tasks, verifies each source file, creates the output directory when needed, loads each workbook with Aspose.Cells, and saves it as a PDF using the task‑specific options while handling exceptions and logging results.
    class ConversionTask
    {
        public string ExcelPath { get; set; } = string.Empty;
        public string PdfOutputPath { get; set; } = string.Empty;
        public PdfSaveOptions Options { get; set; } = new PdfSaveOptions();
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Prepare a list of conversion tasks with optional PDF settings per file
            var tasks = new List<ConversionTask>
            {
                new ConversionTask
                {
                    ExcelPath = @"C:\Input\Report1.xlsx",
                    PdfOutputPath = @"C:\Output\Report1.pdf",
                    Options = new PdfSaveOptions
                    {
                        // Fit the entire sheet on one page
                        OnePagePerSheet = true
                    }
                },
                new ConversionTask
                {
                    ExcelPath = @"C:\Input\Report2.xlsx",
                    PdfOutputPath = @"C:\Output\Report2.pdf",
                    Options = new PdfSaveOptions()
                    // Default PDF options used for this file
                }
                // Add more tasks as needed
            };

            foreach (var task in tasks)
            {
                // Verify source Excel file exists
                if (!File.Exists(task.ExcelPath))
                {
                    Console.WriteLine($"Source file not found: {task.ExcelPath}");
                    continue;
                }

                try
                {
                    // Ensure output directory exists
                    var outputDir = Path.GetDirectoryName(task.PdfOutputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Load the workbook
                    var workbook = new Workbook(task.ExcelPath);

                    // Save as PDF using the specified options
                    workbook.Save(task.PdfOutputPath, task.Options);

                    Console.WriteLine($"Converted '{task.ExcelPath}' to PDF at '{task.PdfOutputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{task.ExcelPath}': {ex.Message}");
                }
            }
        }
    }
}
