// Title: Logging Aspose.Cells HTML export details – source workbook, TableCssId, and destination path – in C#
// AI Prompts: Generate C# code that loads a workbook, saves it as HTML using Aspose.Cells with a custom CSS Id for the table, and writes a UTC‑timestamped entry to a log file containing the source file path, the CSS Id, and the output HTML path. | Write C# error‑handling logic that catches exceptions from the Aspose.Cells HTML save operation and appends the exception message with a timestamp to the same log file. | Create C# logic that checks for the existence of the output directory before saving the HTML file and creates it if missing.
// Common Searches: c# Aspose.Cells export workbook to html with custom TableCssId and audit log | how to record source and destination paths when saving HTML with Aspose.Cells | append export details to a text file during Aspose.Cells HTML conversion | ensure output folder exists before Aspose.Cells HTML save in .NET | log errors from Aspose.Cells HTML export to a file
// Tags: Aspose.Cells HTML export metadata logging | C# workbook to HTML conversion using CSS identifier | Create output directory before Aspose.Cells save | Write export metadata to a text file in .NET | Exception handling for Aspose.Cells HTML save

using System;
using System.IO;
using Aspose.Cells;

namespace MyExportApp
{
    // The ExportLogger class loads a workbook, saves it as HTML with a specified TableCssId, ensures the output directory exists, and appends a UTC‑timestamped entry with source path, CSS Id, and destination path to a log file, while also recording any exceptions that occur.
    public class ExportLogger
    {
        /// <param name="sourcePath">Full path to the source workbook.</param>
        /// <param name="tableCssId">CSS Id to assign to the exported HTML table.</param>
        /// <param name="destPath">Full path where the HTML file will be saved.</param>
        public static void ExportTable(string sourcePath, string tableCssId, string destPath)
        {
            try
            {
                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

                // Load the source workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(sourcePath);

                // Configure HTML save options (including the TableCssId)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export as HTML table (default behavior)
                    TableCssId = tableCssId
                };

                // Ensure destination directory exists
                string destDir = Path.GetDirectoryName(destPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                    Directory.CreateDirectory(destDir);

                // Save the workbook as HTML (lifecycle rule: save)
                workbook.Save(destPath, saveOptions);

                // Prepare audit log entry
                string logEntry = $"Timestamp: {DateTime.UtcNow:O} | SourceWorkbook: {sourcePath} | TableCssId: {tableCssId} | Destination: {destPath}";
                File.AppendAllText("ExportAuditLog.txt", logEntry + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Log any errors to the audit log and rethrow
                string errorEntry = $"Timestamp: {DateTime.UtcNow:O} | Error: {ex.Message}";
                File.AppendAllText("ExportAuditLog.txt", errorEntry + Environment.NewLine);
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage; adjust paths as needed.
            string sourcePath = "Sample.xlsx";
            string tableCssId = "myTable";
            string destPath = "Output.html";

            try
            {
                ExportLogger.ExportTable(sourcePath, tableCssId, destPath);
                Console.WriteLine("Export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Export failed: {ex.Message}");
            }
        }
    }
}
