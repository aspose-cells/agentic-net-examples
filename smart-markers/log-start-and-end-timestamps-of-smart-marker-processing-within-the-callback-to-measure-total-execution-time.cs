// Title: Log start and end timestamps of smart marker processing using ISmartMarkerCallBack in Aspose.Cells for .NET
// AI Prompts: Create an ISmartMarkerCallBack implementation that records the first and last processing timestamps of smart markers and returns the elapsed TimeSpan. | Demonstrate how to assign the timing callback to a WorkbookDesigner, invoke Process, retrieve the total duration, and output it in milliseconds. | Add per‑marker console logging that shows sheet index, row index, column index, table name, and column name while measuring overall processing time.
// Common Searches: aspnet log smart marker processing start and end time with Aspose.Cells | measure execution time of smart markers using ISmartMarkerCallBack | how to get total smart marker processing duration in C# Aspose.Cells | record timestamps for each smart marker callback in WorkbookDesigner | performance monitoring of smart markers Aspose.Cells .NET
// Tags: smart marker timing callback | Aspose.Cells processing duration measurement | ISmartMarkerCallBack timestamp logging | WorkbookDesigner smart marker performance | C# log smart marker execution time

using System;
using System.Data;
using Aspose.Cells;

// The example defines a SmartMarkerTimingCallback that captures the first and last timestamps of smart marker processing, provides a GetDuration method, attaches the callback to a WorkbookDesigner, processes markers with a DataTable source, prints the total processing time in milliseconds, and saves the resulting workbook.
public class SmartMarkerTimingCallback : ISmartMarkerCallBack
{
    private DateTime? _startTime;
    private DateTime? _endTime;

    // Called for each smart marker during processing
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        // Record start time on first call
        if (_startTime == null)
            _startTime = DateTime.Now;

        // Update end time on every call (will hold time of last processed marker)
        _endTime = DateTime.Now;

        // Optional per‑marker logging
        Console.WriteLine($"Processed marker at {DateTime.Now:O} - Sheet:{sheetIndex} Row:{rowIndex} Col:{colIndex} Table:{tableName} Column:{columnName}");
    }

    // Returns the total duration between first and last marker processing
    public TimeSpan? GetDuration()
    {
        if (_startTime.HasValue && _endTime.HasValue)
            return _endTime - _startTime;
        return null;
    }
}

public class Program
{
    public static void Main()
    {
        // Load the template workbook that contains smart markers
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = new Workbook("SmartMarkerTemplate.xlsx");

        // Attach the callback that measures processing time
        SmartMarkerTimingCallback callback = new SmartMarkerTimingCallback();
        designer.CallBack = callback;

        // Prepare a simple data source
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Rows.Add("John", 30);
        dt.Rows.Add("Jane", 28);
        designer.SetDataSource(dt);

        // Process all smart markers in the workbook
        designer.Process(true);

        // Output total processing time measured by the callback
        TimeSpan? duration = callback.GetDuration();
        if (duration.HasValue)
            Console.WriteLine($"Total smart marker processing time: {duration.Value.TotalMilliseconds} ms");
        else
            Console.WriteLine("Processing time could not be determined.");

        // Save the resulting workbook
        designer.Workbook.Save("SmartMarkerResult.xlsx");
    }
}
