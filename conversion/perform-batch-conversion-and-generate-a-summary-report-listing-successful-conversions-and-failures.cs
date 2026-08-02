// Title: C# Batch Spreadsheet Conversion with Aspose.Cells and Summary Report
// Description: A C# sample that processes a list of source‑destination file pairs, checks each source, converts workbooks using Aspose.Cells.Utility.ConversionUtility to formats such as PDF, XLSX, or CSV, records success or error, and writes a comma‑separated report (ConversionReport.txt) summarizing the operation.
// Keywords: Aspose.Cells | C# batch conversion | ConversionUtility | Excel to PDF | XLS to XLSX | CSV to XLSX | conversion log | summary report | file conversion automation | error handling
// Common Searches: Aspose.Cells batch convert Excel files C# | How to generate a conversion log with Aspose.Cells | Convert multiple spreadsheets to PDF using .NET | Record failures during Aspose.Cells conversion | Create a summary report after batch workbook conversion
// Developer Intent: Convert a collection of spreadsheet files to various target formats and produce a concise report that indicates which conversions succeeded and which failed.
// Use Cases: Automate the conversion of monthly financial statements from .xlsx to PDF while keeping an audit trail. | Migrate legacy .xls workbooks to .xlsx and capture files that cannot be processed. | Transform a batch of .csv data exports into .xlsx worksheets and generate a status report for each file.
// AI Prompts: Rewrite the example to output the conversion summary as JSON instead of CSV. | Add parallel processing with Task.WhenAll while preserving per‑file error handling and report generation. | Show how to apply custom PDF export options (e.g., page size, image quality) for each job in the batch.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Cells.Utility;

namespace BatchConversionExample
{
    // A C# sample that processes a list of source‑destination file pairs, checks each source, converts workbooks using Aspose.Cells.Utility.ConversionUtility to formats such as PDF, XLSX, or CSV, records success or error, and writes a comma‑separated report (ConversionReport.txt) summarizing the operation.
    public class Converter
    {
        // Entry point for the batch conversion process
        public static void Run()
        {
            // Define the batch of files to convert (source -> destination)
            var conversionJobs = new List<(string source, string destination)>
            {
                ("InputFiles/Report1.xlsx", "OutputFiles/Report1.pdf"),
                ("InputFiles/Data1.xls", "OutputFiles/Data1.xlsx"),
                ("InputFiles/Chart.csv", "OutputFiles/Chart.xlsx"),
                // Add more jobs as needed
            };

            // Prepare a StringBuilder to collect the summary report
            var reportBuilder = new StringBuilder();
            reportBuilder.AppendLine("Source File,Destination File,Status,Message");

            foreach (var job in conversionJobs)
            {
                try
                {
                    // Ensure the source file exists before attempting conversion
                    if (!File.Exists(job.source))
                        throw new FileNotFoundException("Source file not found.", job.source);

                    // Perform the conversion using Aspose.Cells.Utility.ConversionUtility
                    ConversionUtility.Convert(job.source, job.destination);

                    // Record successful conversion
                    reportBuilder.AppendLine($"{job.source},{job.destination},Success,");
                }
                catch (Exception ex)
                {
                    // Record failure with the exception message
                    string safeMessage = ex.Message.Replace("\"", "\"\"");
                    reportBuilder.AppendLine($"{job.source},{job.destination},Failure,\"{safeMessage}\"");
                }
            }

            // Write the summary report to a text file
            string reportPath = "ConversionReport.txt";
            try
            {
                File.WriteAllText(reportPath, reportBuilder.ToString());
                Console.WriteLine($"Conversion summary report generated at: {Path.GetFullPath(reportPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write the summary report: {ex.Message}");
            }
        }
    }

    // Program entry point required for compilation
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                Converter.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
