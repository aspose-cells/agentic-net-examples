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
            // Input Excel file path (modify as needed)
            string inputPath = "input.xlsx";

            // Output log file path
            string logPath = "PivotRefreshDates.txt";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure all pivot tables are refreshed so RefreshDate is up‑to‑date
            workbook.Worksheets.RefreshPivotTables();

            // Open a StreamWriter for logging
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the collection of pivot tables in the current worksheet
                    PivotTableCollection pivots = sheet.PivotTables;

                    // Iterate through each pivot table
                    foreach (PivotTable pt in pivots)
                    {
                        // Write worksheet name, pivot table name and its RefreshDate
                        writer.WriteLine($"Worksheet: {sheet.Name}, PivotTable: {pt.Name}, RefreshDate: {pt.RefreshDate}");
                    }
                }
            }

            // Optionally save the workbook if any changes were made
            // workbook.Save("output.xlsx");
        }
    }
}