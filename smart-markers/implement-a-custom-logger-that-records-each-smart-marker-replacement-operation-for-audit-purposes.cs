// Title: Create a custom ISmartMarkerCallBack logger to audit smart marker replacements in Aspose.Cells for .NET
// Description: This example shows how to implement a custom logger by inheriting ISmartMarkerCallBack. The logger records the sheet index, row, column, table name, and column name for every smart marker processed, stores entries in a list, writes them to a text file, and integrates with WorkbookDesigner.CallBack before calling Process(). The sample includes error handling, file existence checks, and demonstrates saving the processed workbook and audit log.
// Keywords: Aspose.Cells smart marker callback | ISmartMarkerCallBack logger | audit smart marker replacements | C# Aspose.Cells example | custom smart marker logger .NET | record smart marker processing | save smart marker audit log | WorkbookDesigner CallBack | Aspose.Cells GitHub sample
// Common Searches: how to log smart marker replacements in Aspose.Cells | custom ISmartMarkerCallBack implementation C# | save smart marker processing details to a file | Aspose.Cells audit log for smart markers | example of WorkbookDesigner.CallBack usage
// Developer Intent: Implement a logger that captures each smart marker replacement and outputs an audit file.
// Use Cases: Track which cells are populated by smart markers for compliance reporting. | Debug complex smart marker templates by reviewing detailed replacement logs. | Integrate audit data with external systems such as databases or monitoring tools. | Persist processing information for post‑run analysis in large workbooks.
// AI Prompts: Generate a C# example that extends ISmartMarkerCallBack to write smart marker audit entries to a CSV file. | Show how to store smart marker processing logs in a SQL Server table using Aspose.Cells. | Explain best practices for error handling in a custom smart marker logger for large Excel files.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// This example shows how to implement a custom logger by inheriting ISmartMarkerCallBack. The logger records the sheet index, row, column, table name, and column name for every smart marker processed, stores entries in a list, writes them to a text file, and integrates with WorkbookDesigner.CallBack before calling Process(). The sample includes error handling, file existence checks, and demonstrates saving the processed workbook and audit log.
public class SmartMarkerAuditLogger : ISmartMarkerCallBack
{
    // Stores each smart marker processing event
    public List<string> AuditEntries { get; } = new List<string>();

    // Called by Aspose.Cells for every smart marker replacement
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        string entry = $"Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}";
        AuditEntries.Add(entry);
        Console.WriteLine("SmartMarker processed: " + entry);
    }

    // Writes the collected audit information to a text file
    public void SaveLog(string filePath)
    {
        File.WriteAllLines(filePath, AuditEntries);
    }
}

public class SmartMarkerAuditExample
{
    public static void Run()
    {
        try
        {
            const string templatePath = "TemplateWithSmartMarkers.xlsx";

            // Verify template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load a workbook that contains smart markers
            Workbook template = new Workbook(templatePath);

            // Prepare a simple data source
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // Initialize WorkbookDesigner and assign the workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = template
            };

            // Set the data source for smart markers
            designer.SetDataSource(dt);

            // Create and assign the custom logger as the callback
            SmartMarkerAuditLogger logger = new SmartMarkerAuditLogger();
            designer.CallBack = logger;

            // Process all smart markers in the workbook
            designer.Process();

            // Save the processed workbook
            const string outputPath = "ProcessedOutput.xlsx";
            designer.Workbook.Save(outputPath);
            Console.WriteLine($"Processed workbook saved to: {outputPath}");

            // Persist the audit log
            const string logPath = "SmartMarkerAuditLog.txt";
            logger.SaveLog(logPath);
            Console.WriteLine($"Audit log saved to: {logPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during processing:");
            Console.WriteLine(ex.Message);
        }
    }
}

public class Program
{
    public static void Main()
    {
        SmartMarkerAuditExample.Run();
    }
}
