// Title: Record PivotTable RefreshDate for Every Worksheet with Aspose.Cells (.NET)
// Description: Loads an Excel workbook using Aspose.Cells, iterates through all worksheets and their PivotTables, refreshes each table to update its RefreshDate, and writes the worksheet name, pivot name, and RefreshDate to a text log before saving the file.
// Keywords: Aspose.Cells C# PivotTable RefreshDate | log pivot refresh timestamp | iterate worksheets Aspose.Cells | export pivot refresh date to file | Excel automation audit log | C# write pivot data to text file | Aspose.Cells example RefreshDate
// Common Searches: How to get RefreshDate of each PivotTable with Aspose.Cells | C# code to log pivot table refresh dates | Iterate all worksheets and write pivot timestamps to a file | Aspose.Cells retrieve pivot refresh timestamp | Export PivotTable RefreshDate to text file in .NET
// Developer Intent: Extract the RefreshDate of every PivotTable in a workbook and persist it to a log file.
// Use Cases: Create an audit trail of when each pivot table was last refreshed for compliance reporting. | Validate data freshness by comparing pivot refresh timestamps before publishing a workbook. | Automate monitoring of pivot table updates across multiple sheets in a batch processing pipeline.
// AI Prompts: Generate C# code with Aspose.Cells that lists all PivotTables in a workbook and outputs their RefreshDate without altering the file. | Show how to write PivotTable RefreshDate values to a CSV file with robust error handling and proper resource disposal. | Explain the steps to refresh PivotTables, capture their RefreshDate, and keep the original workbook formatting intact using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel workbook using Aspose.Cells, iterates through all worksheets and their PivotTables, refreshes each table to update its RefreshDate, and writes the worksheet name, pivot name, and RefreshDate to a text log before saving the file.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Path to the log file where refresh dates will be written
        string logFilePath = "PivotRefreshLog.txt";

        // Load the workbook (uses Aspose.Cells load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Open a StreamWriter to create/overwrite the log file
        using (StreamWriter writer = new StreamWriter(logFilePath, false))
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables in the current worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the pivot table to ensure RefreshDate is up‑to‑date
                    pivotTable.RefreshData();
                    pivotTable.CalculateData();

                    // Log worksheet name, pivot table name and its RefreshDate
                    writer.WriteLine($"Worksheet: {sheet.Name}, PivotTable: {pivotTable.Name}, RefreshDate: {pivotTable.RefreshDate}");
                }
            }
        }

        // Save the workbook if any modifications were made (uses Aspose.Cells save rule)
        workbook.Save("output.xlsx");
    }
}
