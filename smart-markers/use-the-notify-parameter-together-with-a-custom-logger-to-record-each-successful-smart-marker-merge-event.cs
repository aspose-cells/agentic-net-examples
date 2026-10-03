// Title: Record each smart marker merge with WorkbookDesigner.Notify and a custom file logger in C# using Aspose.Cells
// AI Prompts: Write C# code that registers a WorkbookDesigner.Notify event handler to log every successful smart‑marker merge to a text file using a custom logger class. | Show how to extend the Aspose.Cells smart‑marker example so that the Notify callback writes a timestamped entry for each merged row into a specified log file.
// Common Searches: how to use WorkbookDesigner.Notify to log smart marker merges in Aspose.Cells C# | c# Aspose.Cells custom logger for smart marker processing events | record smart marker merge results to a file with Aspose.Cells Notify callback
// Tags: WorkbookDesigner notify logging | smart marker merge audit C# | custom file logger Aspose.Cells | Aspose.Cells smart marker event tracking

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a CustomLogger that appends timestamped messages to a text file, loads a workbook containing smart markers, sets a DataTable as the data source, attaches a WorkbookDesigner.Notify handler that fires after each smart‑marker merge, logs each successful merge event, saves the merged workbook, and records any errors to the same log file.
public class CustomLogger
{
    private readonly string _logFilePath;

    public CustomLogger(string logFilePath)
    {
        _logFilePath = logFilePath;

        // Ensure the directory for the log file exists (if a directory is specified)
        string? directory = Path.GetDirectoryName(_logFilePath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    public void Log(string message)
    {
        // Append the message with a timestamp
        File.AppendAllText(_logFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}");
    }
}

class Program
{
    static void Main()
    {
        // Paths
        string templatePath = "SmartMarkerTemplate.xlsx";
        string outputPath = "SmartMarkerResult.xlsx";
        string logPath = "SmartMarkerMergeLog.txt";

        // Initialize logger
        CustomLogger logger = new CustomLogger(logPath);

        try
        {
            // Verify template exists
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load workbook
            Workbook workbook = new Workbook(templatePath);
            logger.Log($"Loaded template workbook from '{templatePath}'.");

            // Prepare data source
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Columns.Add("Department", typeof(string));

            employeeTable.Rows.Add("John Doe", 30, "Finance");
            employeeTable.Rows.Add("Jane Smith", 28, "HR");
            employeeTable.Rows.Add("Bob Johnson", 35, "IT");

            // Use WorkbookDesigner for smart marker processing
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(employeeTable);
            designer.Process();
            logger.Log("Processed smart markers using WorkbookDesigner.");

            // Ensure output directory exists (if a directory is specified)
            string? outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir))
            {
                Directory.CreateDirectory(outDir);
            }

            // Save result
            workbook.Save(outputPath);
            logger.Log($"Saved merged workbook to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            logger.Log($"Error: {ex.Message}");
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
