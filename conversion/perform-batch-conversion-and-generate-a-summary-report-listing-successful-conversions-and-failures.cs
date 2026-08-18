// Title: Batch convert Excel files with Aspose.Cells in C# and generate a conversion summary report
// Description: C# code that scans a folder for Excel‑compatible files, uses Aspose.Cells.Utility.ConversionUtility to convert each file to a target format (e.g., PDF, DOCX), captures successes and errors, and writes a timestamped text report summarizing the batch operation.
// Keywords: Aspose.Cells | C# batch conversion | Excel to PDF .NET | ConversionUtility | folder processing | conversion report | error handling | automate Excel conversion | multiple file conversion | Aspose.Cells example
// Common Searches: Aspose.Cells batch convert Excel to PDF C# | How to create a conversion report with Aspose.Cells | Convert all Excel files in a folder using Aspose.Cells | C# code for bulk Excel to DOCX conversion | Log failed Excel conversions Aspose.Cells
// Developer Intent: Convert every Excel‑related file in a directory to a specified format and produce a detailed log of successful and failed conversions.
// Use Cases: Archive a large set of financial workbooks by converting them to PDF and obtaining a report that flags any files that could not be processed. | Standardize incoming CSV, ODS, or TSV data by converting them to XLSX while capturing conversion errors for audit trails. | Run a nightly job that transforms uploaded Excel workbooks to PDF, stores the files, and generates a summary report for operations monitoring.
// AI Prompts: Write C# code that uses Aspose.Cells ConversionUtility to batch convert Excel files to DOCX and outputs a JSON summary of successes and failures. | Explain how to modify the BatchConverter to execute conversions in parallel threads and add per‑file timestamps to the report. | Show how to add an email notification that sends the generated ConversionReport.txt after the batch conversion finishes.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // C# code that scans a folder for Excel‑compatible files, uses Aspose.Cells.Utility.ConversionUtility to convert each file to a target format (e.g., PDF, DOCX), captures successes and errors, and writes a timestamped text report summarizing the batch operation.
    public class BatchConverter
    {
        // Runs batch conversion of Excel files to the specified format and creates a summary report.
        public void Run(string inputFolder, string outputFolder, string targetExtension)
        {
            // Ensure input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Define supported Excel source extensions
            string[] supportedExtensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".csv", ".ods", ".tsv" };

            // Gather source files
            var sourceFiles = new List<string>();
            foreach (var ext in supportedExtensions)
            {
                sourceFiles.AddRange(Directory.GetFiles(inputFolder, "*" + ext, SearchOption.TopDirectoryOnly));
            }

            var successful = new List<string>();
            var failed = new List<string>();

            foreach (var srcPath in sourceFiles)
            {
                // Verify source file exists (safety check)
                if (!File.Exists(srcPath))
                {
                    failed.Add($"{srcPath} => File not found");
                    continue;
                }

                string destFileName = Path.GetFileNameWithoutExtension(srcPath) + targetExtension;
                string destPath = Path.Combine(outputFolder, destFileName);

                try
                {
                    // Perform the conversion using Aspose.Cells.Utility.ConversionUtility
                    ConversionUtility.Convert(srcPath, destPath);
                    successful.Add(destPath);
                }
                catch (Exception ex)
                {
                    failed.Add($"{srcPath} => {ex.Message}");
                }
            }

            // Build the summary report
            var sb = new StringBuilder();
            sb.AppendLine("Batch Conversion Report");
            sb.AppendLine($"Timestamp: {DateTime.Now}");
            sb.AppendLine($"Total files processed: {sourceFiles.Count}");
            sb.AppendLine($"Successful conversions: {successful.Count}");
            foreach (var ok in successful)
            {
                sb.AppendLine($"  OK: {ok}");
            }
            sb.AppendLine($"Failed conversions: {failed.Count}");
            foreach (var err in failed)
            {
                sb.AppendLine($"  FAIL: {err}");
            }

            // Write the report to a text file
            string reportPath = Path.Combine(outputFolder, "ConversionReport.txt");
            try
            {
                File.WriteAllText(reportPath, sb.ToString());
                Console.WriteLine($"Conversion completed. Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write report: {ex.Message}");
            }
        }
    }

    internal class Program
    {
        // Entry point required for compilation
        private static void Main(string[] args)
        {
            try
            {
                // Example usage – adjust paths as needed
                string inputFolder = @"C:\InputExcel";
                string outputFolder = @"C:\ConvertedFiles";
                string targetExtension = ".pdf";

                var converter = new BatchConverter();
                converter.Run(inputFolder, outputFolder, targetExtension);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
