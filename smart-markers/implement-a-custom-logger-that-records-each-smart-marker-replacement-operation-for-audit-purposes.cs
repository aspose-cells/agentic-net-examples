// Title: Add a custom audit logger to track Aspose.Cells smart marker processing in C#
// AI Prompts: Create a C# AuditLogger class that initializes a log file, ensures the directory exists, and writes timestamped entries for each step of WorkbookDesigner smart‑marker processing. | Show how to integrate the AuditLogger into a workflow that loads an Excel workbook, assigns a DataTable as the data source, calls WorkbookDesigner.Process(), and logs load, data‑source assignment, processing, and save actions. | Demonstrate extending the logger to capture the name of each smart marker and the corresponding data row values during processing.
// Common Searches: c# Aspose.Cells audit log for smart marker replacements | example of logging WorkbookDesigner.Process steps in .NET | how to record each smart marker replacement in an Excel workbook using Aspose.Cells | timestamped audit trail for Excel smart markers with Aspose.Cells | track data source assignment and smart marker processing in C# Aspose.Cells
// Tags: Aspose.Cells smart marker audit logging | WorkbookDesigner processing log implementation | timestamped Excel smart marker replacement tracking | C# audit trail for data source assignment in Aspose.Cells | log file management for Aspose.Cells workbook designer

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample shows how to build an AuditLogger that creates a log file, ensures its folder exists, and writes timestamped entries. It demonstrates loading an Excel workbook, assigning a DataTable as the data source, processing all smart markers with WorkbookDesigner, and logging each major step—load, data source assignment, processing, and save—providing an audit trail for smart‑marker replacements.
class AuditLogger
{
    private readonly string _logPath;

    public AuditLogger(string logPath)
    {
        _logPath = logPath;

        // Ensure directory exists
        var dir = Path.GetDirectoryName(_logPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        // Initialize log file
        File.WriteAllText(_logPath, $"Audit Log started at {DateTime.Now}{Environment.NewLine}");
    }

    public void Log(string message)
    {
        // Append a timestamped entry
        File.AppendAllText(_logPath, $"{DateTime.Now}: {message}{Environment.NewLine}");
    }
}

class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputWithSmartMarkers.xlsx";
            const string outputFile = "OutputProcessed.xlsx";
            const string logFile = "SmartMarkerAudit.log";

            // Verify input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Load workbook containing smart markers
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Initialize the custom audit logger
            var logger = new AuditLogger(logFile);
            logger.Log("Workbook loaded successfully.");

            // Set up WorkbookDesigner to process smart markers
            var designer = new WorkbookDesigner(workbook);

            // Example data source for smart markers
            var dataTable = new DataTable();
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Amount", typeof(double));
            dataTable.Rows.Add("Alice", 123.45);
            dataTable.Rows.Add("Bob", 678.90);

            // Assign the data source
            designer.SetDataSource(dataTable);
            logger.Log("Data source assigned.");

            // Process all smart markers in the workbook
            try
            {
                designer.Process();
                logger.Log("Smart markers processed.");
            }
            catch (Exception ex)
            {
                logger.Log($"Error during smart marker processing: {ex.Message}");
                Console.WriteLine($"Processing error: {ex.Message}");
                return;
            }

            // Ensure output directory exists
            var outDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            // Save the processed workbook
            workbook.Save(outputFile);
            logger.Log($"Workbook saved to '{outputFile}'.");
            Console.WriteLine($"Processing completed. Output saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
