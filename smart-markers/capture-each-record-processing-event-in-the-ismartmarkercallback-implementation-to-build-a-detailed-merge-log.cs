// Title: How to capture and log each smart marker merge event using ISmartMarkerCallBack in Aspose.Cells for .NET
// AI Prompts: Implement an ISmartMarkerCallBack class that appends sheet index, row index, column index, table name, and column name to a List<string> for every smart marker processed. | Demonstrate assigning the custom callback to a WorkbookDesigner, binding a DataTable as the data source, invoking Process(), and printing the accumulated log entries to the console. | Extend the callback so that each merge log entry is written directly to a text file while still maintaining the in‑memory list.
// Common Searches: aspnet log smart marker merge details with ISmartMarkerCallBack | example of capturing smart marker processing events in Aspose.Cells C# | how to output smart marker merge positions to console using WorkbookDesigner | record table and column names for each smart marker during merge Aspose.Cells | save smart marker merge log to file in C# Aspose.Cells
// Tags: smart marker merge logging Aspose.Cells | ISmartMarkerCallBack event capture C# | WorkbookDesigner callback for merge audit | log smart marker cell positions Excel | record smart marker table column details

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerLogging
{
    // Custom callback that records each smart marker processing event
    // The example defines a MergeLogCallback that implements ISmartMarkerCallBack to record sheet, row, column, table, and column identifiers for every smart marker occurrence. The callback is assigned to WorkbookDesigner, a DataTable is set as the data source, Process() populates the workbook, and the collected log entries are printed, illustrating how to audit smart marker merges in Aspose.Cells for .NET.
    public class MergeLogCallback : ISmartMarkerCallBack
    {
        // Stores log entries
        public List<string> LogEntries { get; } = new List<string>();

        // Called by Aspose.Cells for each smart marker occurrence
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Build a detailed log entry
            string entry = $"Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:\"{tableName}\", Column:\"{columnName}\"";
            LogEntries.Add(entry);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (template) and add smart markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            // Smart markers using table "Employees"
            sheet.Cells["A1"].PutValue("&=Employees.Name");
            sheet.Cells["B1"].PutValue("&=Employees.Age");
            sheet.Cells["A2"].PutValue("&=Employees.Name");
            sheet.Cells["B2"].PutValue("&=Employees.Age");

            // Initialize WorkbookDesigner with the template workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Create and assign the custom callback
            MergeLogCallback callback = new MergeLogCallback();
            designer.CallBack = callback;

            // Prepare a simple data source
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Rows.Add("Alice", 30);
            employeeTable.Rows.Add("Bob", 25);
            employeeTable.Rows.Add("Charlie", 35);

            // Set the data source for the designer
            designer.SetDataSource(employeeTable);

            // Process smart markers (populate data)
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("SmartMarkerResult.xlsx");

            // Output the merge log
            Console.WriteLine("Smart Marker Processing Log:");
            foreach (string log in callback.LogEntries)
            {
                Console.WriteLine(log);
            }
        }
    }
}
