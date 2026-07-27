// Title: Audit Aspose.Cells HTML Export – Log Source Workbook, TableCssId, and Destination Path (C#)
// Description: Demonstrates how to add an audit trail to Aspose.Cells HTML conversions. A static logger writes a UTC timestamp, the source workbook location (or a placeholder for a new workbook), the HtmlSaveOptions.TableCssId value, and the full destination file path to a text file before the workbook is saved as HTML.
// Keywords: Aspose.Cells | HTML export audit | log TableCssId | export destination path | C# workbook logging | .NET Excel to HTML | audit trail Excel conversion
// Common Searches: how to log Aspose.Cells HTML export in C# | track source workbook and TableCssId when saving to HTML | audit Excel to HTML conversion Aspose.Cells | record export details for Aspose.Cells HTML save | Aspose.Cells export logging example
// Developer Intent: Implement a simple audit mechanism that records the source workbook, the TableCssId used, and the target HTML file path for every Aspose.Cells HTML export.
// Use Cases: Maintain compliance logs for financial reports exported from Excel to HTML. | Analyze styling consistency by tracking TableCssId values across multiple exports. | Monitor per‑user conversion activity in a multi‑tenant SaaS platform for troubleshooting and reporting.
// AI Prompts: Create C# code that extends ExportAuditLogger to capture export duration and success status for each Aspose.Cells HTML save. | Replace the text‑file logger with Serilog while preserving source workbook, TableCssId, and destination path information. | Write a unit test that verifies ExportAuditLogger writes the correct entry for both newly created and existing workbook exports.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExportAudit
{
    // Simple logger that writes audit information to a text file.
    // Demonstrates how to add an audit trail to Aspose.Cells HTML conversions. A static logger writes a UTC timestamp, the source workbook location (or a placeholder for a new workbook), the HtmlSaveOptions.TableCssId value, and the full destination file path to a text file before the workbook is saved as HTML.
    public static class ExportAuditLogger
    {
        private static readonly string LogFilePath = "ExportAuditLog.txt";

        public static void Log(string sourceWorkbook, string tableCssId, string destinationPath)
        {
            string logEntry = $"Timestamp: {DateTime.UtcNow:u} | Source: {sourceWorkbook} | TableCssId: {tableCssId} | Destination: {destinationPath}";
            File.AppendAllLines(LogFilePath, new[] { logEntry });
        }
    }

    public class WorkbookExporter
    {
        // Exports a workbook to HTML while logging the operation details.
        // sourcePath   : Path to the source Excel file (can be null for a newly created workbook).
        // destPath     : Path where the HTML file will be saved.
        // tableCssId   : Value to assign to HtmlSaveOptions.TableCssId.
        public static void ExportToHtml(string sourcePath, string destPath, string tableCssId)
        {
            Workbook workbook;

            // Load or create the workbook.
            if (!string.IsNullOrEmpty(sourcePath) && File.Exists(sourcePath))
            {
                // Load existing workbook.
                workbook = new Workbook(sourcePath);
            }
            else
            {
                // Create a new workbook with sample data.
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                sheet.Cells["A2"].PutValue("Alice");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Bob");
                sheet.Cells["B3"].PutValue(25);
            }

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                TableCssId = tableCssId,
                ExportWorksheetCSSSeparately = false // default, kept for clarity
            };

            // Log the export operation before saving.
            string sourceInfo = string.IsNullOrEmpty(sourcePath) ? "NewlyCreatedWorkbook" : Path.GetFullPath(sourcePath);
            ExportAuditLogger.Log(sourceInfo, tableCssId, Path.GetFullPath(destPath));

            // Perform the export.
            workbook.Save(destPath, saveOptions);
        }
    }

    class Program
    {
        static void Main()
        {
            // Example 1: Export a newly created workbook.
            WorkbookExporter.ExportToHtml(
                sourcePath: null,
                destPath: "NewWorkbookExport.html",
                tableCssId: "myTableStyle");

            // Example 2: Export an existing workbook.
            string existingExcel = "SampleData.xlsx"; // Ensure this file exists in the execution folder.
            WorkbookExporter.ExportToHtml(
                sourcePath: existingExcel,
                destPath: "ExistingWorkbookExport.html",
                tableCssId: "existingTable");

            Console.WriteLine("Export operations completed. Audit log written to ExportAuditLog.txt");
        }
    }
}
