// Title: C# batch conversion of Excel, CSV, ODS, and other spreadsheets to PDF with Aspose.Cells and CSV summary report
// Common Searches: C# Aspose.Cells batch convert multiple Excel files to PDF and log results to CSV | How to process all .xls, .xlsx, .csv, .ods files in a folder and save as PDF using Aspose.Cells | Generate conversion report for spreadsheet‑to‑PDF batch job in .NET | Aspose.Cells error handling for batch workbook to PDF conversion in C# | Create CSV summary of successful and failed spreadsheet conversions with Aspose.Cells
// Tags: batch spreadsheet to PDF conversion Aspose.Cells | Aspose.Cells save workbook as PDF C# | CSV conversion summary report .NET | error handling during workbook conversion Aspose.Cells | process multiple spreadsheet formats programmatically

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchConversion
{
    // The C# program scans a given input directory for supported spreadsheet files (.xls, .xlsx, .xlsb, .xlsm, .csv, .ods), loads each workbook with Aspose.Cells, saves it as a PDF in an output folder, records success or failure for each file, and writes a CSV report containing source path, destination PDF path, conversion status, and any error messages.
    class Program
    {
        // Structure to hold conversion result for each file
        class ConversionResult
        {
            public string SourceFile { get; set; }
            public string DestinationFile { get; set; }
            public bool Success { get; set; }
            public string Message { get; set; }
        }

        static void Main(string[] args)
        {
            // Input folder containing source spreadsheets
            string inputFolder = @"C:\InputFiles";
            // Output folder where converted PDFs will be placed
            string outputFolder = @"C:\ConvertedFiles";
            // Path for the summary report (CSV format)
            string reportPath = @"C:\ConversionReport.csv";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure output folder exists
            Directory.CreateDirectory(outputFolder);

            // Collect all supported spreadsheet files in the input folder
            string[] supportedExtensions = new[] { ".xls", ".xlsx", ".xlsb", ".xlsm", ".csv", ".ods" };
            var sourceFiles = new List<string>();
            foreach (var ext in supportedExtensions)
            {
                try
                {
                    sourceFiles.AddRange(Directory.GetFiles(inputFolder, "*" + ext, SearchOption.TopDirectoryOnly));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error enumerating files with extension '{ext}': {ex.Message}");
                }
            }

            var results = new List<ConversionResult>();

            foreach (var sourcePath in sourceFiles)
            {
                var result = new ConversionResult
                {
                    SourceFile = sourcePath,
                    DestinationFile = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(sourcePath) + ".pdf")
                };

                try
                {
                    // Verify the source file still exists before loading
                    if (!File.Exists(sourcePath))
                        throw new FileNotFoundException("Source file not found.", sourcePath);

                    // Load the workbook using Aspose.Cells
                    Workbook workbook = new Workbook(sourcePath);

                    // Save as PDF
                    workbook.Save(result.DestinationFile, SaveFormat.Pdf);

                    result.Success = true;
                    result.Message = "Converted successfully.";
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = ex.Message;
                }

                results.Add(result);
            }

            // Generate CSV summary report
            try
            {
                using (var writer = new StreamWriter(reportPath))
                {
                    // Header
                    writer.WriteLine("SourceFile,DestinationFile,Success,Message");

                    foreach (var r in results)
                    {
                        // Escape commas and quotes in paths/messages
                        string sourceEscaped = $"\"{r.SourceFile}\"";
                        string destEscaped = $"\"{r.DestinationFile}\"";
                        string messageEscaped = $"\"{r.Message?.Replace("\"", "\"\"")}\"";

                        writer.WriteLine($"{sourceEscaped},{destEscaped},{r.Success},{messageEscaped}");
                    }
                }

                Console.WriteLine($"Batch conversion completed. Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write report: {ex.Message}");
            }
        }
    }
}
