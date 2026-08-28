// Title: Implement a custom ISmartMarkerCallBack logger for auditing smart marker replacements in Aspose.Cells .NET
// AI Prompts: Write a C# class that implements Aspose.Cells.ISmartMarkerCallBack and stores sheet, row, column, table and column names for every smart marker replacement. | Demonstrate assigning the logger to WorkbookDesigner, invoking Process(), and printing the collected audit entries. | Extend the logger to write each audit entry to a text file instead of the console.
// Common Searches: how to track smart marker replacements with Aspose.Cells .NET callback | example of smart marker callback interface for logging in C# | audit smart marker processing using Aspose.Cells WorkbookDesigner | save smart marker replacement details to a file with Aspose.Cells | custom logger for smart markers Aspose.Cells C# tutorial
// Tags: smart marker callback audit .NET | record smart marker replacements Excel | WorkbookDesigner custom callback C# | Aspose.Cells logging smart markers | audit trail for smart markers

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerLoggingDemo
{
    // Custom logger implementing the ISmartMarkerCallBack interface.
    // It records each smart marker processing event.
    // The example defines a SmartMarkerLogger class that implements ISmartMarkerCallBack, captures sheet index, row index, column index, table name, and column name for each smart marker replacement, stores them in a list, and prints the log. It shows how to load a template workbook, set a DataTable as the data source, assign the logger to WorkbookDesigner, process all smart markers, display the audit log, and save the resulting workbook.
    public class SmartMarkerLogger : ISmartMarkerCallBack
    {
        // Collection to store log entries.
        public List<string> LogEntries { get; } = new List<string>();

        // This method is called by Aspose.Cells for each smart marker replacement.
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Build a log entry with all relevant details.
            string entry = $"Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:\"{tableName}\", Column:\"{columnName}\"";
            LogEntries.Add(entry);

            // Optional: also write to console for immediate feedback.
            Console.WriteLine("SmartMarker processed: " + entry);
        }

        // Helper to output the entire log (e.g., after processing is complete).
        public void PrintLog()
        {
            Console.WriteLine("\n--- Smart Marker Processing Log ---");
            foreach (var entry in LogEntries)
            {
                Console.WriteLine(entry);
            }
            Console.WriteLine("--- End of Log ---\n");
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers.
            // Replace "template.xlsx" with the actual path to your template file.
            Workbook templateWorkbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Prepare a sample data source (DataTable) matching the smart markers in the template.
            DataTable data = new DataTable("Employees");
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Age", typeof(int));
            data.Columns.Add("Department", typeof(string));

            data.Rows.Add("John Doe", 30, "Sales");
            data.Rows.Add("Jane Smith", 28, "Marketing");

            // Set the data source for the designer.
            designer.SetDataSource(data);

            // Create and assign the custom logger as the callback.
            SmartMarkerLogger logger = new SmartMarkerLogger();
            designer.CallBack = logger;

            // Process all smart markers in the workbook.
            // Using the parameterless Process() method processes the entire workbook.
            designer.Process();

            // After processing, output the collected log entries.
            logger.PrintLog();

            // Save the resulting workbook.
            // Replace "output.xlsx" with the desired output path.
            designer.Workbook.Save("output.xlsx");
        }
    }
}
