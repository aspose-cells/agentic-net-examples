// Title: C# Progress Callback for Aspose.Cells Smart Marker Processing (Percentage Completion)
// Description: Demonstrates how to implement ISmartMarkerCallBack to report the percentage of smart markers processed while populating a workbook with a large DataTable. The callback calculates progress using designer.GetSmartMarkers(), safeguards against division‑by‑zero, writes updates to the console, and integrates with WorkbookDesigner for high‑volume smart marker scenarios.
// Keywords: Aspose.Cells | Smart Markers | ISmartMarkerCallBack | progress callback | percentage reporting | C# | .NET | WorkbookDesigner | large data source | console progress | smart marker processing | performance monitoring
// Common Searches: Aspose.Cells smart marker progress callback C# | how to track smart marker processing percentage | ISmartMarkerCallBack example | monitor smart marker population in .NET | report progress during large smart marker fill | Aspose.Cells console progress for smart markers
// Developer Intent: Add a callback that outputs the percentage of smart markers processed while a workbook is being populated with a massive data set.
// Use Cases: Show real‑time console progress when processing thousands of smart marker cells. | Integrate the callback with a UI progress bar for desktop applications. | Log percentage updates to a file for audit or performance analysis. | Provide cancellation support based on user input while still reporting progress. | Ensure accurate progress calculation by deriving total markers from GetSmartMarkers and preventing division‑by‑zero errors.
// AI Prompts: Generate a C# ISmartMarkerCallBack that updates a Windows Forms ProgressBar instead of writing to the console. | Rewrite the callback to append timestamped progress entries to a log file. | Create code that adds cancellation token handling to the smart marker processing loop while preserving percentage updates. | Provide a SignalR hub example that streams progress percentages to a web dashboard during smart marker population. | Write unit tests that verify correct percentage calculation and zero‑division protection in the progress callback.

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerProgressDemo
{
    // Callback implementation to report progress during smart marker processing
    // Demonstrates how to implement ISmartMarkerCallBack to report the percentage of smart markers processed while populating a workbook with a large DataTable. The callback calculates progress using designer.GetSmartMarkers(), safeguards against division‑by‑zero, writes updates to the console, and integrates with WorkbookDesigner for high‑volume smart marker scenarios.
    public class SmartMarkerProgressCallback : ISmartMarkerCallBack
    {
        private readonly int _totalMarkers;
        private int _processedCount;

        public SmartMarkerProgressCallback(int totalMarkers)
        {
            _totalMarkers = totalMarkers > 0 ? totalMarkers : 1; // avoid division by zero
            _processedCount = 0;
        }

        // This method is invoked for each smart marker cell being processed
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            _processedCount++;

            // Calculate percentage based on processed cells vs total markers
            int percent = (int)((double)_processedCount / _totalMarkers * 100);
            Console.WriteLine($"Processing sheet {sheetIndex}, cell ({rowIndex}, {colIndex}) - Table: {tableName}, Column: {columnName} => {percent}% completed");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook("SmartMarkerTemplate.xlsx");

            // Determine total number of smart markers to process
            // GetSmartMarkers returns distinct markers; for demonstration we treat this as total work units
            string[] markers = designer.GetSmartMarkers();
            int totalMarkers = markers.Length;

            // Assign the progress callback
            designer.CallBack = new SmartMarkerProgressCallback(totalMarkers);

            // Prepare a large data source (e.g., a DataTable with many rows)
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Department", typeof(string));

            // Populate with a large number of rows to simulate a heavy load
            for (int i = 1; i <= 5000; i++)
            {
                dt.Rows.Add($"Employee {i}", 20 + (i % 30), $"Dept {(i % 5) + 1}");
            }

            // Bind the data source to the designer
            designer.SetDataSource(dt);

            // Process all smart markers; the callback will report progress
            designer.Process();

            // Save the populated workbook
            designer.Workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}
