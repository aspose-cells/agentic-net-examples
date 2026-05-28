using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotRefreshDateLogger
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be processed
            string inputPath = "input.xlsx";

            // Path to the log file where RefreshDate information will be written
            string logPath = "PivotRefreshDates.log";

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Open a StreamWriter for logging (overwrites existing file)
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the collection of pivot tables in the current worksheet
                    PivotTableCollection pivotTables = sheet.PivotTables;

                    // If there are no pivot tables, continue to next worksheet
                    if (pivotTables == null || pivotTables.Count == 0)
                        continue;

                    // Iterate through each pivot table
                    foreach (PivotTable pivot in pivotTables)
                    {
                        // Retrieve the RefreshDate property
                        DateTime refreshDate = pivot.RefreshDate;

                        // Write worksheet name, pivot table name, and refresh date to the log
                        writer.WriteLine($"Worksheet: {sheet.Name}, PivotTable: {pivot.Name}, RefreshDate: {refreshDate}");
                    }
                }
            }

            // Optionally save the workbook (save rule) – unchanged in this scenario
            workbook.Save("output.xlsx");
        }
    }
}