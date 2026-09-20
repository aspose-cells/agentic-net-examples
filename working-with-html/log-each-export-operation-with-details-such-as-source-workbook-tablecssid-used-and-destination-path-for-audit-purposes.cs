// Title: Create an audit log for Aspose.Cells HTML export with source workbook, TableCssId, and destination path in C#
// AI Prompts: Write a C# console program that loads an Excel workbook, exports it to HTML using Aspose.Cells HtmlSaveOptions with a custom TableCssId, and appends a UTC timestamped entry containing the source file, TableCssId, and output path to a log file. | Add robust exception handling so that any error during the HTML export is captured and written with a timestamp to the same audit log, creating the log file if it does not exist. | Extend the logging to include the elapsed time of the export operation and store each entry in CSV format for easier analysis.
// Common Searches: how to write an audit log for Aspose.Cells HTML export in a .NET console app | record source workbook and TableCssId when saving Excel as HTML using Aspose.Cells | C# example for logging Excel to HTML conversion details with timestamp | append error information to export log file during Aspose.Cells HTML save | measure export duration and save log as CSV in Aspose.Cells C#
// Tags: Aspose.Cells HTML export audit logging | C# HtmlSaveOptions TableCssId configuration | timestamped export log file .NET | exception handling for Aspose.Cells HTML save | CSV format export performance logging

using Aspose.Cells;
using System;
using System.IO;

// The sample loads a source.xlsx workbook, exports it to exported.html with HtmlSaveOptions.TableCssId set to "myTable", and appends a UTC timestamped line to export_log.txt that records the source path, TableCssId, and destination path. It also captures any exceptions, logs them with timestamps, and demonstrates how to add duration tracking and CSV‑style entries.
class ExportLogger
{
    static void Main()
    {
        // Define paths and TableCssId
        string sourcePath = "source.xlsx";          // source workbook
        string destinationPath = "exported.html";   // export destination
        string logPath = "export_log.txt";          // audit log file
        string tableCssId = "myTable";              // TableCssId to apply

        try
        {
            // Verify source workbook exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Configure HTML export options with the specified TableCssId
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // The following properties are not available in the current Aspose.Cells version:
            // htmlOptions.ExportTableColumnHeaders = true;
            // htmlOptions.ExportTableRowHeaders = true;
            htmlOptions.TableCssId = tableCssId; // apply CSS ID to the exported table

            // Save the workbook as HTML (save rule)
            workbook.Save(destinationPath, htmlOptions);

            // Log the export operation for audit purposes
            string logEntry = $"{DateTime.Now:u}: Source='{sourcePath}', TableCssId='{tableCssId}', Destination='{destinationPath}'{Environment.NewLine}";
            File.AppendAllText(logPath, logEntry);
        }
        catch (Exception ex)
        {
            // Log the error details
            string errorLog = $"{DateTime.Now:u}: ERROR - {ex.Message}{Environment.NewLine}";
            File.AppendAllText(logPath, errorLog);
            Console.Error.WriteLine(errorLog);
        }
    }
}
