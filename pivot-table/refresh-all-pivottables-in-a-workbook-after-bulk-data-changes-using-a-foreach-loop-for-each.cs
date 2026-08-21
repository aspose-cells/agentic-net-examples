// Title: C# – Refresh All PivotTables in an Aspose.Cells Workbook After Bulk Data Changes
// Description: Loads a workbook, modifies source cells, iterates through each worksheet with a foreach loop, calls RefreshPivotTables to update every pivot table, and saves the result. Demonstrates error handling and best practices for bulk data updates.
// Keywords: Aspose.Cells | RefreshPivotTables | C# pivot table refresh | bulk data update | foreach loop Excel | update all pivot tables | Excel automation | global | US
// Common Searches: how to refresh all pivot tables in Aspose.Cells C# | foreach loop refresh pivot tables workbook | bulk data change refresh pivot tables Aspose | programmatically update pivot tables after data edit | Aspose.Cells refresh all pivots example
// Developer Intent: Programmatically refresh every PivotTable in a workbook after making bulk data modifications.
// Use Cases: Automated monthly reports where source data is altered and all pivot tables must reflect the new values before export. | Data import routines that change many cells and require immediate pivot table updates across multiple worksheets. | Scheduled Excel processing jobs that need to ensure pivot tables stay synchronized with dynamic data sources.
// AI Prompts: Write C# code using Aspose.Cells to change a range of cells and then refresh all pivot tables in each worksheet. | Show how to add robust exception handling around RefreshPivotTables calls after bulk data updates. | Provide an example that refreshes pivot tables only in worksheets whose names start with "Sales" using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, modifies source cells, iterates through each worksheet with a foreach loop, calls RefreshPivotTables to update every pivot table, and saves the result. Demonstrates error handling and best practices for bulk data updates.
    public class RefreshAllPivotTablesDemo
    {
        public static void Run()
        {
            string inputPath = "InputData.xlsx";
            string outputPath = "RefreshedPivotTables.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(inputPath);

                // Example bulk data changes
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Cells["B2"].PutValue(1500);
                dataSheet.Cells["B3"].PutValue(2500);
                dataSheet.Cells["B4"].PutValue(3500);

                // Refresh all PivotTables in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.RefreshPivotTables();
                }

                // Save updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshAllPivotTablesDemo.Run();
        }
    }
}
