using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// Custom logger implementing ISmartMarkerCallBack to capture each smart marker replacement
public class SmartMarkerAuditLogger : ISmartMarkerCallBack
{
    private readonly List<string> _entries = new List<string>();

    // This method is invoked by Aspose.Cells for every smart marker replacement
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        string entry = $"Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}";
        _entries.Add(entry);
        Console.WriteLine("SmartMarker processed: " + entry);
    }

    // Persist the collected log entries to a text file
    public void SaveLog(string filePath)
    {
        File.WriteAllLines(filePath, _entries);
    }
}

public class SmartMarkerAuditDemo
{
    public static void Main()
    {
        // Load a template workbook that contains smart markers
        Workbook template = new Workbook("TemplateWithSmartMarkers.xlsx");

        // Prepare a simple data source
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Rows.Add("John Doe", 30);
        dt.Rows.Add("Jane Smith", 28);

        // Initialize WorkbookDesigner and assign the template workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = template;

        // Set the data source for smart markers
        designer.SetDataSource(dt);

        // Attach the custom logger
        SmartMarkerAuditLogger logger = new SmartMarkerAuditLogger();
        designer.CallBack = logger;

        // Process all smart markers in the workbook
        designer.Process();

        // Save the processed workbook
        designer.Workbook.Save("ProcessedOutput.xlsx");

        // Save the audit log to a file
        logger.SaveLog("SmartMarkerAuditLog.txt");
    }
}