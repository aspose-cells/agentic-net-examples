// Title: How to audit Aspose.Cells HTML export in C# by logging source workbook, TableCssId, and destination path
// Common Searches: C# Aspose.Cells export to HTML with audit log of source workbook and destination file | log TableCssId when saving an Excel workbook as HTML using Aspose.Cells .NET | record export details to a text file during Aspose.Cells HTML conversion in C# | append workbook export information to a log file with Aspose.Cells HTMLSaveOptions | track HTML export operations for Excel workbooks using Aspose.Cells in .NET
// Tags: Aspose.Cells HTML export logging | C# workbook to HTML audit trail | HtmlSaveOptions TableCssId usage | export details text file append | Aspose.Cells export audit log implementation

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions to apply a custom TableCssId, ensures the output directory exists, saves the workbook as HTML, and then writes a UTC‑timestamped entry containing the source workbook path, TableCssId, and destination HTML path to an ExportAuditLog.txt file in the same folder. Errors are caught and written to the console.
public class WorkbookExporter
{
    /// <param name="sourceWorkbookPath">Full path of the source workbook.</param>
    /// <param name="tableCssId">CSS Id to be applied to the exported HTML table.</param>
    /// <param name="destinationHtmlPath">Full path where the HTML file will be saved.</param>
    public static void ExportToHtml(string sourceWorkbookPath, string tableCssId, string destinationHtmlPath)
    {
        try
        {
            // Verify source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(sourceWorkbookPath))
                throw new FileNotFoundException($"Source workbook not found: {sourceWorkbookPath}");

            // Load the source workbook.
            Workbook workbook = new Workbook(sourceWorkbookPath);

            // Configure HTML save options with the provided TableCssId.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                TableCssId = tableCssId
            };

            // Ensure the destination directory exists.
            string destDir = Path.GetDirectoryName(destinationHtmlPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(destDir))
                Directory.CreateDirectory(destDir);

            // Save the workbook as HTML.
            workbook.Save(destinationHtmlPath, htmlOptions);

            // Prepare log entry.
            string logEntry = $"{DateTime.UtcNow:O} | SourceWorkbook: {sourceWorkbookPath} | TableCssId: {tableCssId} | DestinationHtml: {destinationHtmlPath}";

            // Determine log file location (same folder as the destination HTML).
            string logFilePath = Path.Combine(destDir, "ExportAuditLog.txt");

            // Append the log entry to the audit log file.
            File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
        }
        catch (Exception ex)
        {
            // Log the exception details to the console (or a separate error log if desired).
            Console.Error.WriteLine($"Error exporting workbook to HTML: {ex.Message}");
            // Optionally rethrow or handle further.
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Example usage: provide paths via command‑line arguments or use defaults.
        string sourcePath = args.Length > 0 ? args[0] : @"C:\Data\Sample.xlsx";
        string cssId = args.Length > 1 ? args[1] : "myTable";
        string destPath = args.Length > 2 ? args[2] : @"C:\Exports\Sample.html";

        try
        {
            WorkbookExporter.ExportToHtml(sourcePath, cssId, destPath);
            Console.WriteLine("Export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
