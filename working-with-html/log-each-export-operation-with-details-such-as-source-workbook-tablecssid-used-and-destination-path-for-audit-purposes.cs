// Title: C# – Log Excel‑to‑HTML Export with TableCssId and Destination Path using Aspose.Cells
// Description: Shows how to audit each HTML export performed with Aspose.Cells for .NET. The ExportLogger loads a workbook, configures HtmlSaveOptions with a custom TableCssId and an IExportObjectListener, saves to HTML, and writes the source file name, output path, TableCssId, and details of every exported object (type and index) to the console or any logging sink.
// Keywords: Aspose.Cells | C# HTML export | ExportObjectListener | TableCssId | Excel to HTML audit | export logging .NET | track exported images | Aspose.Cells HtmlSaveOptions | audit conversion | log workbook export
// Common Searches: Aspose.Cells log HTML export C# | how to use ExportObjectListener with HtmlSaveOptions | record TableCssId during Excel to HTML conversion | audit Excel to HTML conversion Aspose.Cells | log each exported image shape Aspose.Cells | save workbook as HTML with custom table id | track export operations Aspose.Cells .NET
// Developer Intent: I need to record details of every Excel‑to‑HTML conversion, including source file, TableCssId, output path, and each exported object.
// Use Cases: Compliance reporting for batch Excel‑to‑HTML conversions | Debugging missing images or charts after HTML export | Integrating export logs into a monitoring dashboard | Storing export metadata in a database for later analysis | Generating CSV audit trails for regulatory purposes
// AI Prompts: Write code to persist ExportLogger entries to a CSV file with timestamps and operation IDs. | Create a unit test that asserts CustomExportObjectListener logs the correct object type and sequential index. | Show how to modify ExportLogger to write logs to an Azure Table storage. | Provide a PowerShell script that parses the console output and inserts records into a SQL Server table. | Generate a sample GitHub Actions workflow that runs the export logger on multiple workbooks and uploads the log as an artifact.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Logs export operations and handles object export events
// Shows how to audit each HTML export performed with Aspose.Cells for .NET. The ExportLogger loads a workbook, configures HtmlSaveOptions with a custom TableCssId and an IExportObjectListener, saves to HTML, and writes the source file name, output path, TableCssId, and details of every exported object (type and index) to the console or any logging sink.
class ExportLogger
{
    private readonly string _sourcePath;
    private readonly string _destPath;
    private readonly string _tableCssId;

    public ExportLogger(string sourcePath, string destPath, string tableCssId)
    {
        _sourcePath = sourcePath;
        _destPath = destPath;
        _tableCssId = tableCssId;
    }

    public void ExportToHtml()
    {
        // Load the source workbook
        Workbook workbook = new Workbook(_sourcePath);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            TableCssId = _tableCssId,
            ExportObjectListener = new CustomExportObjectListener(_sourcePath, _tableCssId)
        };

        // Save the workbook as HTML
        workbook.Save(_destPath, saveOptions);

        // Log the overall export operation
        LogExport();
    }

    private void LogExport()
    {
        Console.WriteLine($"[Export] Workbook: '{Path.GetFileName(_sourcePath)}' => '{_destPath}'");
        Console.WriteLine($"[Export] TableCssId used: '{_tableCssId}'");
        // Extend this method to write to a file, database, etc., if needed
    }
}

// Implements IExportObjectListener to log each exported object (e.g., images, shapes)
class CustomExportObjectListener : IExportObjectListener
{
    private readonly string _sourcePath;
    private readonly string _tableCssId;
    private int _objectIndex = 0;

    public CustomExportObjectListener(string sourcePath, string tableCssId)
    {
        _sourcePath = sourcePath;
        _tableCssId = tableCssId;
    }

    public object ExportObject(ExportObjectEvent e)
    {
        _objectIndex++;
        object source = e?.GetSource();
        string typeName = source?.GetType().Name ?? "null";

        Console.WriteLine($"[ExportObject] #{_objectIndex}: Type={typeName}, Workbook='{Path.GetFileName(_sourcePath)}', TableCssId='{_tableCssId}'");

        // Return null to let Aspose.Cells perform the default export handling
        return null;
    }
}

// Demonstrates usage
class Program
{
    static void Main()
    {
        string sourceFile = "input.xlsx";          // Path to the source workbook
        string destinationFile = "output.html";    // Desired HTML output path
        string tableCssId = "myTableCss";          // Custom TableCssId for the HTML table

        ExportLogger exporter = new ExportLogger(sourceFile, destinationFile, tableCssId);
        exporter.ExportToHtml();
    }
}
